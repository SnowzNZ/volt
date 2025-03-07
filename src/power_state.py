from enum import Enum


class PowerState(Enum):
    UNKNOWN = ("unknown", "Unknown")
    BATTERY = ("on_battery", "On Battery")
    AC = ("plugged_in", "Plugged In")

    def __init__(self, raw_name, display_name):
        self.raw_name = raw_name
        self.display_name = display_name
