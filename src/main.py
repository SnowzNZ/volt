import subprocess
import uuid

from PIL import Image
from pystray import Icon, Menu, MenuItem


class Volt:
    def __init__(self, icon_path="icon.png"):
        self.icon_path = icon_path
        self.icon = None

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

    def create_menu_item(self, guid: uuid.UUID, name: str, is_active: bool):
        def on_click():
            self.set_power_plan(guid)

        return MenuItem(name, on_click, checked=lambda item: is_active)

    def update_menu(self):
        power_plans, active_guid = self.get_power_plans()

        menu_items = [
            self.create_menu_item(guid, name, guid == active_guid)
            for guid, name in power_plans
        ]
        menu_items.append(MenuItem("Exit", lambda: self.stop()))

        self.icon.menu = Menu(*menu_items)

    def initialize_tray(self) -> None:
        power_plans, active_guid = self.get_power_plans()

        menu_items = [
            self.create_menu_item(guid, name, guid == active_guid)
            for guid, name in power_plans
        ]
        menu_items.append(MenuItem("Exit", lambda: self.stop()))

        menu = Menu(*menu_items)
        self.icon = Icon("volt", Image.open(self.icon_path), "Volt", menu)
        self.icon.run()

    def stop(self):
        self.icon.stop()


if __name__ == "__main__":
    volt_app = Volt()
    volt_app.initialize_tray()
