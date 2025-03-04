import win32com.client as win32
import time

CANoe = win32.DispatchEx("CANoe.Application")

CANoe.Open(r"C:\Users\srava\Desktop\LED\LED_Blinking\Indicator.cfg")
time.sleep(3)
CANoe.measurement.Start()
time.sleep(3)

#system variables
systemCAN = CANoe.System.namespaces
sys_namespace = systemCAN("Indicator")
sys_value = sys_namespace.Variables("Sys.BCM")

sys_value.Value = 1 #system variable value
time.sleep(0.5)
print(sys_value.Value)
#print(sys_value)

sys_sig_value = CANoe.GetBus("CAN").GetSignal(1, "BCMMessage", "BCM_Siganl") #signal value
print("BCM_Signal value:", sys_sig_value)

CANoe.measurement.Stop()
time.sleep(3)
CANoe.Quit()


