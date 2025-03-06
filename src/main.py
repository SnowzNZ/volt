import json
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


class Volt:
    def __init__(self, icon_path="icon.png"):
        self.icon_path = icon_path
        self.icon = None
        self.power_state: PowerState = PowerState.UNKNOWN
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
            print(f"Error running powercfg: {e}")
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
                        print(f"Invalid GUID: {guid_part}")

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
                        print(f"Invalid GUID: {guid_part}")

        return power_plans, active_guid

    def set_power_plan(self, guid: uuid.UUID) -> None:
        try:
            subprocess.run(["powercfg", "/S", str(guid)], check=True)
            self.update_menu()
        except subprocess.CalledProcessError as e:
            print(f"Error setting power plan: {e}")

    def create_menu_item(
        self, guid: uuid.UUID, name: str, is_active: bool, power_state: PowerState
    ):
        def on_click():
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
                print(f"Invalid GUID in saved plans: {guid}")

    def update_menu(self):
        power_plans, active_guid = self.get_power_plans()

        menu_items = [
            MenuItem(
                "Plugged In",
                Menu(
                    *[
                        self.create_menu_item(
                            guid, name, guid == active_guid, PowerState.AC
                        )
                        for guid, name in power_plans
                    ]
                ),
            ),
            MenuItem(
                "On Battery",
                Menu(
                    *[
                        self.create_menu_item(
                            guid, name, guid == active_guid, PowerState.BATTERY
                        )
                        for guid, name in power_plans
                    ]
                ),
            ),
            MenuItem("Exit", lambda: self.stop()),
        ]

        self.icon.menu = Menu(*menu_items)

    def initialize_tray(self) -> None:
        power_plans, active_guid = self.get_power_plans()

        menu_items = [
            MenuItem(
                "Plugged In",
                Menu(
                    *[
                        self.create_menu_item(
                            guid, name, guid == active_guid, PowerState.AC
                        )
                        for guid, name in power_plans
                    ],
                ),
            ),
            MenuItem(
                "On Battery",
                Menu(
                    *[
                        self.create_menu_item(
                            guid, name, guid == active_guid, PowerState.BATTERY
                        )
                        for guid, name in power_plans
                    ],
                ),
            ),
            MenuItem("Exit", lambda: self.stop()),
        ]

        menu = Menu(*menu_items)
        self.icon = Icon("volt", Image.open(self.icon_path), "Volt", menu)
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
                    self.apply_saved_plan("on_battery")
                elif battery_status["ACLineStatus"] == 1:
                    self.power_state = PowerState.AC
                    self.apply_saved_plan("plugged_in")
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
