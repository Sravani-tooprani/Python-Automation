import win32com.client as win32
import time
from report_generator import *
from DiagReqRes import *

#open the CANoe application
CANoe = win32.DispatchEx("CANoe.Application")

def acc_active():
    # file = open(r"C:\Desktop\ADAS_Automation\ADAS_TestCases\Report.txt",'w')
    write_content("ACC Test Cases\n")
    write_content("TestCase1: ACC Active\n")

    write_content("TestStep1: Set Mo_Sig_ReadytoDrive to 0x2 (drive)\n")
    Mo_drive_sig = CANoe.GetBus("CAN").GetSignal(1, "MO_Message_1", "MO_sig_ReadytoDrive")
    time.sleep(0.1)
    Mo_drive_sig.Value = 0x2
    time.sleep(0.5)

    write_content("TestStep2: Set ACC_Int_status to 0x1 (active)\n")
    ACC_init_status = CANoe.GetBus("CAN").GetSignal(1, "ACC_BCM_Message", "ACC_Int_status")
    time.sleep(0.1)
    ACC_init_status.Value = 0x1
    time.sleep(0.5)

    write_content("TestStep3: Set ACC_MainSwitch_ACC to 0x2 (active)\n")
    ACC_Main_Switch = CANoe.GetBus("CAN").GetSignal(1, "ACC_BCM_Message", "ACC_MainSwitch_ACC")
    time.sleep(0.1)
    ACC_Main_Switch.Value = 0x2
    time.sleep(0.5)

    write_content("TestStep4: Set Gear_signal to 0x3 (drive)\n")
    Gear_Signal = CANoe.GetBus("CAN").GetSignal(1, "Gear_Message", "Gear_signal")
    time.sleep(0.1)
    Gear_Signal.Value = 0x3
    time.sleep(0.5)

    write_content("TestStep5: Set Sig_Spdmtr_Rq > 40kmph\n")
    Speed_sig = CANoe.GetBus("CAN").GetSignal(1, "RDS_Coder_Data_CAN", "Sig_Spdmtr_Rq")
    time.sleep(0.1)
    Speed_sig.Value = 0x78
    time.sleep(0.5)

    write_content("TestStep6: Set ACC_Set_signal to 0x1 (active)\n")
    ACC_Set_signal = CANoe.GetBus("CAN").GetSignal(1, "ACC_BCM_Message", "ACC_Set_signal")
    time.sleep(0.1)
    ACC_Set_signal.Value = 0x1
    time.sleep(0.5)

    write_content("TestStep7: Expected: Check ACC_Status should be active\n")
    acc_status = CANoe.GetBus("CAN").GetSignal(1, "ACC_status", "ACC_sig_status")

    if int(acc_status) == 0x2:
        write_content(f"TestStep7 Passed: ACC status is 0x2 Active, measured value is {acc_status}\n")
    else:
        write_content(f"TestStep7 Failed: ACC status is not equal to 0x2 Active, measured value is {acc_status}\n")

def acc_passive():
    # file = open(r"C:\Desktop\ADAS_Automation\ADAS_TestCases\Report.txt", 'a')
    write_content("TestCase2: ACC Passive\n")

    write_content("TestStep1: Set Mo_Sig_ReadytoDrive to 0x2 (drive)\n")
    Mo_drive_sig = CANoe.GetBus("CAN").GetSignal(1, "MO_Message_1", "MO_sig_ReadytoDrive")
    time.sleep(0.1)
    Mo_drive_sig.Value = 0x2
    time.sleep(0.5)

    write_content("TestStep2: Set ACC_Int_status to 0x1 (active)\n")
    ACC_init_status = CANoe.GetBus("CAN").GetSignal(1, "ACC_BCM_Message", "ACC_Int_status")
    time.sleep(0.1)
    ACC_init_status.Value = 0x1
    time.sleep(0.5)

    write_content("TestStep3: Set ACC_MainSwitch_ACC to 0x1 (OFF)\n")
    ACC_Main_Switch = CANoe.GetBus("CAN").GetSignal(1, "ACC_BCM_Message", "ACC_MainSwitch_ACC")
    time.sleep(0.1)
    ACC_Main_Switch.Value = 0x2
    time.sleep(0.5)

    write_content("TestStep4: Set Gear_signal to 0x3 (drive)\n")
    Gear_Signal = CANoe.GetBus("CAN").GetSignal(1, "Gear_Message", "Gear_signal")
    time.sleep(0.1)
    Gear_Signal.Value = 0x3
    time.sleep(0.5)

    write_content("TestStep5: Set Sig_Spdmtr_Rq > 40kmph\n")
    Speed_sig = CANoe.GetBus("CAN").GetSignal(1, "RDS_Coder_Data_CAN", "Sig_Spdmtr_Rq")
    time.sleep(0.1)
    Speed_sig.Value = 0x78
    time.sleep(0.5)

    write_content("TestStep6: Set ACC_Set_signal to 0x1 (active)\n")
    ACC_Set_signal = CANoe.GetBus("CAN").GetSignal(1, "ACC_BCM_Message", "ACC_Set_signal")
    time.sleep(0.1)
    ACC_Set_signal.Value = 0x1
    time.sleep(0.5)

    write_content("TestStep7: Expected: Check ACC_Status should be passive\n")
    acc_status = CANoe.GetBus("CAN").GetSignal(1, "ACC_status", "ACC_sig_status")

    if int(acc_status) == 0x2:
        write_content(f"TestStep7 Passed: ACC status is 0x1 passive, measured value is {acc_status}\n")
    else:
        write_content(f"TestStep7 Failed: ACC status is not equal to 0x1 passive, measured value is {acc_status}\n")


def acc_passive_apc_speed_condition():
    write_content("TestCase3: ACC Passive and APC Speed Condition\n")

    write_content("TestStep1: Set Mo_Sig_ReadytoDrive to 0x2 (drive)\n")
    Mo_drive_sig = CANoe.GetBus("CAN").GetSignal(1, "MO_Message_1", "MO_sig_ReadytoDrive")
    time.sleep(0.1)
    Mo_drive_sig.Value = 0x2
    time.sleep(0.5)

    write_content("TestStep2: Set ACC_Int_status to 0x1 (active)\n")
    ACC_init_status = CANoe.GetBus("CAN").GetSignal(1, "ACC_BCM_Message", "ACC_Int_status")
    time.sleep(0.1)
    ACC_init_status.Value = 0x1
    time.sleep(0.5)

    write_content("TestStep3: Set ACC_MainSwitch_ACC to 0x2 (active)\n")
    ACC_Main_Switch = CANoe.GetBus("CAN").GetSignal(1, "ACC_BCM_Message", "ACC_MainSwitch_ACC")
    time.sleep(0.1)
    ACC_Main_Switch.Value = 0x2
    time.sleep(0.5)

    write_content("TestStep4: Set Gear_signal to 0x3 (drive)\n")
    Gear_Signal = CANoe.GetBus("CAN").GetSignal(1, "Gear_Message", "Gear_signal")
    time.sleep(0.1)
    Gear_Signal.Value = 0x3
    time.sleep(0.5)

    write_content("TestStep5: Set Sig_Spdmtr_Rq < 40kmph\n")
    Speed_sig = CANoe.GetBus("CAN").GetSignal(1, "RDS_Coder_Data_CAN", "Sig_Spdmtr_Rq")
    time.sleep(0.1)
    Speed_sig.Value = 0x20
    time.sleep(0.5)

    write_content("TestStep6: Set ACC_Set_signal to 0x1 (active)\n")
    ACC_Set_signal = CANoe.GetBus("CAN").GetSignal(1, "ACC_BCM_Message", "ACC_Set_signal")
    time.sleep(0.1)
    ACC_Set_signal.Value = 0x1
    time.sleep(0.5)

    write_content("TestStep7: Expected - ACC_Status should be passive and vehicle speed is too low\n")
    # Check ACC_sig_status
    acc_sig_status = CANoe.GetBus("CAN").GetSignal(1, "ACC_status", "ACC_sig_status")
    if int(acc_sig_status) == 0x1:
        write_content(f"TestStep7 Passed: ACC_sig_status is 0x1 (Passive), measured value: {acc_sig_status}\n")
    else:
        write_content(f"TestStep7 Failed: Expected ACC_sig_status 0x1 (Passive), but got {acc_sig_status}\n")

    # Check ACC_disp_signal
    acc_disp_signal = CANoe.GetBus("CAN").GetSignal(1, "ACC_status", "ACC_disp_signal")
    if int(acc_disp_signal) == 0x2:
        write_content(f"TestStep7 Passed: ACC_disp_signal is 0x2 (Vehicle speed too low), measured value: {acc_disp_signal}\n")
    else:
        write_content(f"TestStep7 Failed: Expected ACC_disp_signal 0x2 (Vehicle speed too low), but got {acc_disp_signal}\n")


def acc_passive_apc_gear_position():
    write_content("TestCase4: ACC Passive and APC Speed Condition\n")

    write_content("TestStep1: Set Mo_Sig_ReadytoDrive to 0x2 (drive)\n")
    Mo_drive_sig = CANoe.GetBus("CAN").GetSignal(1, "MO_Message_1", "MO_sig_ReadytoDrive")
    time.sleep(0.1)
    Mo_drive_sig.Value = 0x2
    time.sleep(0.5)

    write_content("TestStep2: Set ACC_Int_status to 0x1 (active)\n")
    ACC_init_status = CANoe.GetBus("CAN").GetSignal(1, "ACC_BCM_Message", "ACC_Int_status")
    time.sleep(0.1)
    ACC_init_status.Value = 0x1
    time.sleep(0.5)

    write_content("TestStep3: Set ACC_MainSwitch_ACC to 0x2 (active)\n")
    ACC_Main_Switch = CANoe.GetBus("CAN").GetSignal(1, "ACC_BCM_Message", "ACC_MainSwitch_ACC")
    time.sleep(0.1)
    ACC_Main_Switch.Value = 0x2
    time.sleep(0.5)

    write_content("TestStep4: Set Gear_signal to 0x1 (reverse)\n")
    Gear_Signal = CANoe.GetBus("CAN").GetSignal(1, "Gear_Message", "Gear_signal")
    time.sleep(0.1)
    Gear_Signal.Value = 0x1
    time.sleep(0.5)

    write_content("TestStep5: Set Sig_Spdmtr_Rq > 40kmph\n")
    Speed_sig = CANoe.GetBus("CAN").GetSignal(1, "RDS_Coder_Data_CAN", "Sig_Spdmtr_Rq")
    time.sleep(0.1)
    Speed_sig.Value = 0x78
    time.sleep(0.5)

    write_content("TestStep6: Set ACC_Set_signal to 0x1 (active)\n")
    ACC_Set_signal = CANoe.GetBus("CAN").GetSignal(1, "ACC_BCM_Message", "ACC_Set_signal")
    time.sleep(0.1)
    ACC_Set_signal.Value = 0x1
    time.sleep(0.5)

    write_content("TestStep7: Expected - ACC_Status should be passive and check gear position\n")
    # Check ACC_sig_status
    acc_sig_status = CANoe.GetBus("CAN").GetSignal(1, "ACC_status", "ACC_sig_status")
    if int(acc_sig_status) == 0x2:
        write_content(f"TestStep7 Passed: ACC_sig_status is 0x2 (Passive), measured value: {acc_sig_status}\n")
    else:
        write_content(f"TestStep7 Failed: Expected ACC_sig_status 0x2 (Passive), but got {acc_sig_status}\n")

    # Check ACC_disp_signal
    acc_disp_signal = CANoe.GetBus("CAN").GetSignal(1, "ACC_status", "ACC_disp_signal")
    if int(acc_disp_signal) == 0x3:
        write_content(f"TestStep7 Passed: ACC_disp_signal is 0x3 (check gear position), measured value: {acc_disp_signal}\n")
    else:
        write_content(f"TestStep7 Failed: Expected ACC_disp_signal 0x3 (check gear position), but got {acc_disp_signal}\n")


def acc_active_SOC_vehicle_speed():
    write_content("TestCase5: ACC Active then SOC Condition vehicle speed\n")

    write_content("TestStep1: Set Mo_Sig_ReadytoDrive to 0x2 (drive)\n")
    Mo_drive_sig = CANoe.GetBus("CAN").GetSignal(1, "MO_Message_1", "MO_sig_ReadytoDrive")
    time.sleep(0.1)
    Mo_drive_sig.Value = 0x2
    time.sleep(0.5)

    write_content("TestStep2: Set ACC_Int_status to 0x1 (active)\n")
    ACC_init_status = CANoe.GetBus("CAN").GetSignal(1, "ACC_BCM_Message", "ACC_Int_status")
    time.sleep(0.1)
    ACC_init_status.Value = 0x1
    time.sleep(0.5)

    write_content("TestStep3: Set ACC_MainSwitch_ACC to 0x2 (active)\n")
    ACC_Main_Switch = CANoe.GetBus("CAN").GetSignal(1, "ACC_BCM_Message", "ACC_MainSwitch_ACC")
    time.sleep(0.1)
    ACC_Main_Switch.Value = 0x2
    time.sleep(0.5)

    write_content("TestStep4: Set Gear_signal to 0x3 (drive)\n")
    Gear_Signal = CANoe.GetBus("CAN").GetSignal(1, "Gear_Message", "Gear_signal")
    time.sleep(0.1)
    Gear_Signal.Value = 0x3
    time.sleep(0.5)

    write_content("TestStep5: Set Sig_Spdmtr_Rq > 40kmph\n")
    Speed_sig = CANoe.GetBus("CAN").GetSignal(1, "RDS_Coder_Data_CAN", "Sig_Spdmtr_Rq")
    time.sleep(0.1)
    Speed_sig.Value = 0x78
    time.sleep(0.5)

    write_content("TestStep6: Set ACC_Set_signal to 0x1 (active)\n")
    ACC_Set_signal = CANoe.GetBus("CAN").GetSignal(1, "ACC_BCM_Message", "ACC_Set_signal")
    time.sleep(0.1)
    ACC_Set_signal.Value = 0x1
    time.sleep(0.5)

    write_content("TestStep7: Expected: Check ACC_Status should be active\n")
    acc_status = CANoe.GetBus("CAN").GetSignal(1, "ACC_status", "ACC_sig_status")

    if int(acc_status) == 0x2:
        write_content(f"TestStep7 Passed: ACC status is 0x2 Active, measured value is {acc_status}\n")
    else:
        write_content(f"TestStep7 Failed: ACC status is not equal to 0x2 Active, measured value is {acc_status}\n")

    write_content("TestStep8: Set Sig_Spdmtr_Rq to < 40kmph\n")
    Speed_sig.Value = 0x28
    time.sleep(0.5)

    write_content("TestStep9: Set ACC_Set_signal to 0x1(Active) again\n")
    ACC_Set_signal.Value = 0x1
    time.sleep(0.5)

    write_content("TestStep10: Checking if ACC_sig_status is Passive (0x1) and ACC_disp_signal is Vehicle speed too low (0x2)\n")

    acc_status = CANoe.GetBus("CAN").GetSignal(1, "ACC_status", "ACC_sig_status")
    acc_disp_signal = CANoe.GetBus("CAN").GetSignal(1, "ACC_status", "ACC_disp_signal")

    if int(acc_status) == 0x1:
        write_content(f"TestStep10: Passed: ACC_sig_status is 0x1 (Passive), measured value: {acc_status}\n")
    else:
        write_content(f"TestStep10: Failed: Expected ACC_sig_status 0x1 (Passive), but got {acc_status}\n")

    if int(acc_disp_signal) == 0x2:
        write_content(f"TestStep10: Passed: ACC_disp_signal is 0x2 (Vehicle speed too low), measured value: {acc_disp_signal}\n")
    else:
        write_content(f"TestStep10: Failed: Expected ACC_disp_signal 0x2 (Vehicle speed too low), but got {acc_disp_signal}\n")


def acc_active_SOC_gear_position():
    write_content("TestCase5: ACC Active then SOC Condition Gear Position\n")

    write_content("TestStep1: Setting MO_sig_ReadytoDrive to 0x2 (Drive)\n")
    Mo_drive_sig = CANoe.GetBus("CAN").GetSignal(1, "MO_Message_1", "MO_sig_ReadytoDrive")
    Mo_drive_sig.Value = 0x2
    time.sleep(0.5)

    write_content("TestStep2: Setting ACC_Int_status to 0x1 (Active)\n")
    ACC_init_status = CANoe.GetBus("CAN").GetSignal(1, "ACC_BCM_Message", "ACC_Int_status")
    ACC_init_status.Value = 0x1
    time.sleep(0.5)

    write_content("TestStep3: Setting ACC_MainSwitch_ACC to 0x2 (Active)\n")
    ACC_Main_Switch = CANoe.GetBus("CAN").GetSignal(1, "ACC_BCM_Message", "ACC_MainSwitch_ACC")
    ACC_Main_Switch.Value = 0x2
    time.sleep(0.5)

    write_content("TestStep4: Setting Gear_signal to 0x3 (Drive)\n")
    Gear_Signal = CANoe.GetBus("CAN").GetSignal(1, "Gear_Message", "Gear_signal")
    Gear_Signal.Value = 0x3
    time.sleep(0.5)

    write_content("TestStep5: Setting Sig_Spdmtr_Rq to > 40km/h\n")
    Speed_sig = CANoe.GetBus("CAN").GetSignal(1, "RDS_Coder_Data_CAN", "Sig_Spdmtr_Rq")
    Speed_sig.Value = 0x78
    time.sleep(0.5)

    write_content("TestStep6: Setting ACC_Set_signal to 0x1 (Activated)\n")
    ACC_Set_signal = CANoe.GetBus("CAN").GetSignal(1, "ACC_BCM_Message", "ACC_Set_signal")
    ACC_Set_signal.Value = 0x1
    time.sleep(0.5)

    write_content("TestStep7: Checking if ACC_sig_status is Active (0x2)\n")
    acc_status = CANoe.GetBus("CAN").GetSignal(1, "ACC_status", "ACC_sig_status")

    if int(acc_status) == 0x2:
        write_content(f"TestStep7 Passed: ACC status is 0x2 (Active), measured value: {acc_status}\n")
    else:
        write_content(f"TestStep7 Failed: Expected ACC_status 0x2 (Active), but got {acc_status}\n")

    write_content("TestStep8: Setting Gear_signal to 0x1 (Reverse)\n")
    Gear_Signal.Value = 0x1
    time.sleep(0.5)

    write_content("TestStep9: Setting ACC_Set_signal to 0x1 (Activated) again\n")
    ACC_Set_signal.Value = 0x1
    time.sleep(0.5)

    write_content("TestStep10: Checking if ACC_sig_status is Passive (0x1) and ACC_disp_signal is Check Gear Position (0x3)\n")
    acc_status = CANoe.GetBus("CAN").GetSignal(1, "ACC_status", "ACC_sig_status")
    acc_disp_signal = CANoe.GetBus("CAN").GetSignal(1, "ACC_status", "ACC_disp_signal")

    if int(acc_status) == 0x1:
        write_content(f"TestStep10 Passed: ACC_sig_status is 0x1 (Passive), measured value: {acc_status}\n")
    else:
        write_content(f"TestStep10 Failed: Expected ACC_sig_status 0x1 (Passive), but got {acc_status}\n")

    if int(acc_disp_signal) == 0x3:
        write_content(f"TestStep10 Passed: ACC_disp_signal is 0x3 (Check Gear Position), measured value: {acc_disp_signal}\n")
    else:
        write_content(f"TestStep10 Failed: Expected ACC_disp_signal 0x3 (Check Gear Position), but got {acc_disp_signal}\n")



# CANoe.Open(r"C:\Desktop\ADAS_Automation\RBS_Python_v0.2\RBS_Python.cfg")
# time.sleep(2)
# CANoe.Measurement.Start()
# time.sleep(7)

# call all the functions
diag_req_res()
acc_active()
acc_passive()
acc_passive_apc_speed_condition()
acc_passive_apc_gear_position()
acc_active_SOC_vehicle_speed()
acc_active_SOC_gear_position()

CANoe.Measurement.Stop()
print("Executed Successfully: Reports are in path:C:\Desktop\ADAS_Automation\ADAS_TestCases\Report.html")

