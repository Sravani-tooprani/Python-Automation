import win32com.client as win32
import time

#open the CANoe application
CANoe = win32.DispatchEx("CANoe.Application")
CANoe.Open(r"C:\Desktop\RBS\Configuration1.cfg")

#start the measurement
time.sleep(5)
CANoe.measurement.Start()
time.sleep(5)

#retrieve the signal value
BCM_Signal_Value = CANoe.GetBus("CAN").GetSignal(1, "BCMMessage", "BCM_Signal")
print("BCM_Signal value:", BCM_Signal_Value)

#set signal value which modify in CANoe
BCM_Signal_Value.value = 1
time.sleep(5)
print("BCM_Signal value:", BCM_Signal_Value)

#stop the measurement and close the CANoe
time.sleep(0.5)
CANoe.measurement.Stop()
time.sleep(3)
CANoe.Quit()

