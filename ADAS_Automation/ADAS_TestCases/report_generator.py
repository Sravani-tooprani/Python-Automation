import os

report_title = "Report: ACC TEST REPORT"
file_path = r'C:\Desktop\ADAS_Automation\ADAS_TestCases\Report.html'

initialized = False  # Flag to track if the file has been initialized in the session


def write_content(argument):
    global initialized

    # Check if the file needs initialization (only check and delete if it's the first call in the session)
    if not initialized:
        initialize_file()

    input_string_lower = argument.lower()

    with open(file_path, "a") as file:
        # Check if the report title exists in the file content before appending additional content if not
        # check_report_title(): #this is not required file.write("<h1 style='color: red;text-align:
        # center;background: lightgrey; border-bottom: 2px solid black;border-radius: 8px;'> {} </h1>".format(
        # report_title))

        if "test case" in input_string_lower:
            file.write("<h2 style='color:black;background: yellow'> {} </h2>".format(argument))
        elif "passed" in input_string_lower:
            file.write("<h3 style='color:black;background: green'>{}</h3>".format(argument))
        elif "failed" in input_string_lower:
            file.write("<h3 style='color:black;background: red'>{}</h3>".format(argument))
        else:
            file.write("<h3 style='color:black;background: lightgrey'> {} </h3>".format(argument))

    return file_path  # not required


def initialize_file():
    global initialized

    # Check if the file exists
    if os.path.exists(file_path):
        # Delete the existing file
        os.remove(file_path)

    # Create a new file with the specified report title header
    with open(file_path, "w") as file:
        file.write(
            "<h1 style='color: red;text-align: center;background: lightgrey; border-bottom: 2px solid "
            "black;border-radius: 8px;'> {} </h1>".format(
                report_title))

    initialized = True  # Set the flag to indicate that the file has been initialized

    return file_path  # we can make it none


"""def check_report_title():
    # Check if the report title exists in the file content
    if os.path.exists(file_path):
        with open(file_path, "r") as file:
            for line in file:
                if "<h1" in line and report_title in line:
                    return True
                break  # Stop checking after the first line
    return False"""
