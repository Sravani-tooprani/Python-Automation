# coding: utf-8
"""API for setup/usage of Canoe COM Client interface.
"""
# --------------------------------------------------------------------------
# Standard library imports
import os
import sys
import subprocess
# import win32com.client
import time
import threading
from win32com.client import *
from win32com.client.connect import *



# Vector Canoe Class
class CANoe:
    def __init__(self):
        self.application = None
        # check if there is any instance of CANoe process
        # output = subprocess.check_output('tasklist', shell=True)
        # if CANoe process is still available, kill the process
        # if "CANoe32.exe" in str(output):
        #     os.system("taskkill /im CANoe32.exe /f 2>nul >nul")

        # re-dispatch object for CANoe Application
        self.application = win32com.client.DispatchEx("CANoe.Application")
        self.ver = self.application.Version
        print('Loaded CANoe version ',
            self.ver.major, '.',
            self.ver.minor, '.',
            self.ver.Build, '...')#, sep,''

        self.Measurement = self.application.Measurement.Running
        print(self.Measurement)


    def open_simulation(self, cfgname):
        # open CANoe simulation
        if (self.application != None):
            # check for valid file and it is *.cfg file
            if os.path.isfile(cfgname) and (os.path.splitext(cfgname)[1] == ".cfg"):
                self.application.Open(cfgname)
            else:
                raise RuntimeError("Can't find CANoe cfg file")
        else:
            raise RuntimeError("CANoe Application is missing,unable to open simulation")

    def close_simulation(self):
        # close CANoe simulation
        if (self.application != None):
            self.stop_Measurement()
            self.application.Quit()

        # make sure the CANoe is close properly, otherwise enforce taskkill
        output = subprocess.check_output('tasklist', shell=True)

        if "CANoe32.exe" in str(output):
            os.system("taskkill /im CANoe32.exe /f 2>nul >nul")

        self.application = None

    def start_Measurement(self):
        retry = 0
        retry_counter = 5
        # try to establish measurement within 20s timeout
        while not self.application.Measurement.Running and (retry < retry_counter):
            self.application.Measurement.Start()
            time.sleep(1)
            retry += 1
        if (retry == retry_counter):
            raise RuntimeWarning("CANoe start measuremet failed, Please Check Connection!")

    def stop_Measurement(self):
        if self.application.Measurement.Running:
            self.application.Measurement.Stop()
        else:
            pass

    def get_SysVar(self, ns_name, sysvar_name):

        if (self.application != None):
            systemCAN = self.application.System.Namespaces
            sys_namespace = systemCAN(ns_name)
            sys_value = sys_namespace.Variables(sysvar_name)
            print("Sys_BCM:", sys_value)
            return sys_value.Value

        else:
            raise RuntimeError("CANoe is not open,unable to GetVariable")

    def set_SysVar(self, ns_name, sysvar_name, var):

        if (self.application != None):
            systemCAN = self.application.System.Namespaces
            sys_namespace = systemCAN(ns_name)
            sys_value = sys_namespace.Variables(sysvar_name)
            sys_value.Value = var
            print(sys_value)
            # result = sys_value(sys_name)
            #
            # result = var
        else:
            raise RuntimeError("CANoe is not open,unable to GetVariable")

    def get_SigVal(self, channel_num, msg_name, sig_name, bus_type="CAN"):
        """
        @summary Get the value of a raw CAN signal on the CAN simulation bus
        @param channel_num - Integer value to indicate from which channel we will read the signal, usually start from 1,
                             Check with CANoe can channel setup.
        @param msg_name - String value that indicate the message name to which the signal belong. Check DBC setup.
        @param sig_name - String value of the signal to be read
        @param bus_type - String value of the bus type - e.g. "CAN", "LIN" and etc.
        @return The CAN signal value in floating point value.
                Even if the signal is of integer type, we will still return by
                floating point value.
        @exception None
        """
        if (self.application != None):
            result = self.application.GetBus(bus_type).GetSignal(channel_num, msg_name, sig_name)
            return result.Value
        else:
            raise RuntimeError("CANoe is not open,unable to GetVariable")


ref_var = CANoe()
ref_var.open_simulation(r"C:\Desktop\LED\Configuration1.cfg")
time.sleep(3)
ref_var.start_Measurement()
time.sleep(3)
ref_var.get_SysVar("Indicator", "Sys_BCM")
time.sleep(3)
ref_var.set_SysVar("Indicator","Sys_BCM",1)

ref_var.get_SigVal(1,"BCMMessage","BCM_Signal")

ref_var.stop_Measurement()
time.sleep(3)
ref_var.close_simulation()