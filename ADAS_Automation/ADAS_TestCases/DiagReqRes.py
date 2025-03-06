from py_canoe import CANoe
import time


def diag_req_res():

    diag = CANoe()
    diag.open(r"C:\Desktop\ADAS_Automation\RBS_Python_v0.2\RBS_Python.cfg")
    diag.start_measurement()
    time.sleep(3)

    def_session = diag.send_diag_request(diag_ecu_qualifier_name="Door", request='10 01')
    time.sleep(1)

    ext_session = diag.send_diag_request(diag_ecu_qualifier_name="Door", request='10 03')
    time.sleep(1)

    ACC_var_code = diag.send_diag_request(diag_ecu_qualifier_name="Door", request='2E DE 02 01 02')
    time.sleep(1)
