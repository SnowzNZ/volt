import json
import logging as logger
import os
import subprocess
import threading
import uuid

import win32api
import win32gui
from PIL import Image
from pystray import Icon, Menu, MenuItem

from power_state import PowerState

WM_POWERBROADCAST = 0x0218
PBT_APMPOWERSTATUSCHANGE = 0x000A

CONFIG_FILE = "power_plans.json"
LOGS_DIR = "logs"

os.makedirs(LOGS_DIR, exist_ok=True)

logger.basicConfig(
    filename=os.path.join(LOGS_DIR, "volt.log"),
    level=logger.DEBUG,
    format="%(asctime)s - %(levelname)s - %(message)s",
)


class Volt:
    def __init__(self, icon_path="icon.png"):
        self.icon_path = icon_path
        self.icon = None
        self.power_state: PowerState = self.get_current_power_state()
        self.monitor_thread = None
        self.saved_plans = self.load_saved_plans()

    def load_saved_plans(self):
        if os.path.exists(CONFIG_FILE):
            with open(CONFIG_FILE, "r") as f:
                return json.load(f)
        return {"plugged_in": None, "on_battery": None}

    def save_plans(self):
        with open(CONFIG_FILE, "w") as f:
            json.dump(self.saved_plans, f)

    def get_power_plans(self) -> tuple[list[tuple[uuid.UUID, str]], uuid.UUID]:
        try:
            result = subprocess.run(
                ["powercfg", "/L"], capture_output=True, text=True, check=True
            )
            output = result.stdout
        except subprocess.CalledProcessError as e:
            logger.error(f"Error running powercfg: {e}")
            return [], None

        power_plans: list[tuple[uuid.UUID, str]] = []
        lines = output.splitlines()
        active_guid = None

        for line in lines:
            if "*" in line and "Power Scheme GUID:" in line:
                parts = line.split("(", 1)
                if len(parts) == 2:
                    guid_part = parts[0].split(":")[1].strip()
                    try:
                        active_guid = uuid.UUID(guid_part)
                        break
                    except ValueError:
                        logger.error(f"Invalid GUID: {guid_part}")

        for line in lines:
            if "Power Scheme GUID:" in line:
                parts = line.split("(", 1)
                if len(parts) == 2:
                    guid_part = parts[0].split(":")[1].strip()
                    name_part = parts[1].replace(")", "").replace("*", "").strip()

                    try:
                        guid = uuid.UUID(guid_part)
                        power_plans.append((guid, name_part))
                    except ValueError:
                        logger.error(f"Invalid GUID: {guid_part}")

        return power_plans, active_guid

    def set_power_plan(self, guid: uuid.UUID) -> None:
        try:
            subprocess.run(["powercfg", "/S", str(guid)], check=True)
            self.update_menu()
        except subprocess.CalledProcessError as e:
            logger.error(f"Error setting power plan: {e}")

    def get_current_power_state(self) -> PowerState:
        battery_status = win32api.GetSystemPowerStatus()
        if battery_status["ACLineStatus"] == 1:
            return PowerState.AC
        else:
            return PowerState.BATTERY

    def create_menu_item(
        self, guid: uuid.UUID, name: str, is_active: bool, power_state: PowerState
    ):
        def on_click():
            if power_state == self.power_state:
                self.set_power_plan(guid)
            self.save_plan_for_state(power_state, guid),

        return MenuItem(
            name,
            on_click,
            checked=lambda item: str(guid)
            == self.saved_plans.get(power_state.raw_name),
        )

    def save_plan_for_state(self, power_state: PowerState, guid: uuid.UUID):
        self.saved_plans[power_state.raw_name] = str(guid)
        self.save_plans()

    def apply_saved_plan(self, state: str):
        guid = self.saved_plans.get(state)
        if guid:
            try:
                self.set_power_plan(uuid.UUID(guid))
            except ValueError:
                logger.error(f"Invalid GUID in saved plans: {guid}")

    def generate_power_menu_items(self, active_guid: uuid.UUID):
        return [
            MenuItem(
                "**Plugged In**" if self.power_state == PowerState.AC else "Plugged In",
                Menu(
                    *[
                        self.create_menu_item(
                            guid, name, guid == active_guid, PowerState.AC
                        )
                        for guid, name in self.get_power_plans()[0]
                    ],
                ),
            ),
            MenuItem(
                (
                    "**On Battery**"
                    if self.power_state == PowerState.BATTERY
                    else "On Battery"
                ),
                Menu(
                    *[
                        self.create_menu_item(
                            guid, name, guid == active_guid, PowerState.BATTERY
                        )
                        for guid, name in self.get_power_plans()[0]
                    ],
                ),
            ),
        ]

    def update_menu(self):
        power_plans, active_guid = self.get_power_plans()
        menu_items = self.generate_power_menu_items(active_guid) + [
            MenuItem("Exit volt", lambda: self.icon.stop()),
        ]
        self.icon.menu = Menu(*menu_items)

    def initialize_tray(self) -> None:
        power_plans, active_guid = self.get_power_plans()
        menu_items = self.generate_power_menu_items(active_guid) + [
            MenuItem("Exit volt", lambda: self.icon.stop()),
        ]
        menu = Menu(*menu_items)
        self.icon = Icon("volt", Image.open(self.icon_path), "Volt", menu)
        self.icon.title = f"Volt - {self.power_state.display_name}"
        self.icon.run()

    def stop(self):
        self.icon.stop()
        if self.monitor_thread:
            self.monitor_thread.join()

    def monitor_power_state(self):
        def wndproc(hwnd, msg, wparam, lparam):
            if msg == WM_POWERBROADCAST and wparam == PBT_APMPOWERSTATUSCHANGE:
                battery_status = win32api.GetSystemPowerStatus()
                if battery_status["ACLineStatus"] == 0:
                    self.power_state = PowerState.BATTERY
                elif battery_status["ACLineStatus"] == 1:
                    self.power_state = PowerState.AC

                self.apply_saved_plan(self.power_state)
                self.icon.title = f"Volt - {self.power_state.display_name}"
            return win32gui.DefWindowProc(hwnd, msg, wparam, lparam)

        wc = win32gui.WNDCLASS()
        wc.lpfnWndProc = wndproc
        wc.hInstance = win32api.GetModuleHandle(None)
        wc.lpszClassName = "PowerStatusMonitor"

        win32gui.RegisterClass(wc)

        hwnd = win32gui.CreateWindow(
            wc.lpszClassName,
            "PowerStatusMonitor",
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            wc.hInstance,
            None,
        )

        while True:
            win32gui.PumpWaitingMessages()

    def start_power_monitoring(self):
        self.monitor_thread = threading.Thread(
            target=self.monitor_power_state, daemon=True
        )
        self.monitor_thread.start()


if __name__ == "__main__":
    volt_app = Volt()
    volt_app.start_power_monitoring()
    volt_app.initialize_tray()
