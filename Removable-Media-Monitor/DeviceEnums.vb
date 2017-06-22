Public Enum DeviceTypes
    DeviceInterface = 5
    Handle = 6
    Volume = &H2
End Enum

Public Enum DeviceEvent
    '' system detected a new device
    DeviceArrival = &H8000

    '' Preparing to remove (any program can disable the removal)
    DeviceQueryRemove = &H8001

    '' removed 
    DeviceRemoveComplete = &H8004
End Enum