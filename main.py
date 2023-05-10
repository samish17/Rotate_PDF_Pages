import os
import sys
import PyPDF2
import pyautogui
import easygui
import shutil
from plyer import notification  # pip install pyler, used for notification prompt in windows.

# Asking the user to select the file to rotate.
ask = pyautogui.confirm('Select the pdf file to rotate pages.', buttons=['Select', 'Cancel'])

# Starting the loop for error handling and to rum the program if every thing goes smooth.
while True:

    # Running the program if any error doesn't occur.
    try:
        # proceeding forward if the user selects a file and cancels a process if the user clicks 'cancel'.
        if ask == 'Select':
            # Get the input PDF file location/name
            input_file_location = easygui.fileopenbox(msg='Select a file to rotate the pages.', title="Select File")

            # getting a file name from the location of a file e.g. input_file.pdf
            file_name = os.path.basename(input_file_location)

            # Check if the user clicked 'Cancel' while choosing the input file
            if input_file_location is None:
                pyautogui.confirm('The process has been cancelled.', buttons=['Ok'])
                sys.exit(0)
            else:

                # Open the input PDF file
                with open(input_file_location, "rb") as input_file:

                    # Create a PyPDF2 object
                    pdf_reader = PyPDF2.PdfReader(input_file)
                    # Get the number of pages in the PDF file
                    num_pages = len(pdf_reader.pages)

                    # Get the page numbers that you want to rotate
                    page_numbers = pyautogui.prompt("Enter the page numbers that you want to rotate, separated by "
                                                    "commas: ")

                    # Check if the user clicked cancel
                    if page_numbers is None:
                        pyautogui.confirm('The process has been cancelled.', buttons=['Ok'])
                        sys.exit(0)

                    elif len(page_numbers) == 0:
                        pyautogui.prompt('Please provide the page number.')
                        continue

                    else:

                        # Iterate over the pages in the PDF file
                        for page_num in page_numbers:
                            # Get the page object
                            page = pdf_reader.pages[int(page_num)]
                            break

                        # Rotate the page
                        degree = pyautogui.confirm('Select the degree of rotation.', buttons=['90', '180', '270'])

                        rotation = page.rotate(int(degree))
                        if rotation is None:
                            pyautogui.confirm('The process has been cancelled.', buttons=['Ok'])
                            sys.exit(0)
                        else:

                            # Creating the new name for the edited/rotated file
                            new_file_location = os.path.join(os.path.dirname(input_file_location),
                                                             "rotated_" + file_name)

                            # Create a new PDF file
                            # "wb" mode opens the file in binary format for writing.
                            with open(new_file_location, "wb") as output_file:
                                # Write the rotated pages to the new PDF file
                                pdf_writer = PyPDF2.PdfWriter()

                                for page in pdf_reader.pages:
                                    # pdf_writer.addPage(page)
                                    pdf_writer.add_page(page)

                                pdf_writer.write(output_file)

                            # Get the directory path of the current file
                            save_folder = easygui.diropenbox(msg="Select a folder to save the file", title="Save File")

                            # Check if the user has permission to access the file
                            if not os.access(new_file_location, os.R_OK):
                                # Change the permissions of the file
                                os.chmod(new_file_location, 0o644)

                            # Moving the created file from the base location (location where the program is saved) to
                            # the location of user's choice.
                            shutil.move(new_file_location, save_folder)  # <--- Using Shutil Module

                            print("The PDF file has been rotated.")


                            def notify_task_completed():
                                """This function notify the user when the task is completed and opens the file in
                                file manager."""
                                # Set the title and message of the notification
                                title = 'Task Completed.'
                                message = 'File is successfully extracted to {}'.format(new_file_location)
                                # Display the notification
                                notification.notify(
                                    title=title,
                                    message=message,
                                    app_name='PDF Rotator',
                                    timeout=30
                                )
                                # Opening the file in the file manager after the task is completed.
                                os.startfile(save_folder)


                            notify_task_completed()  # <----- Calling the function

                            sys.exit(0)

        # Terminating the program, if the user clicks cancel while selecting the input file.
        elif ask == 'Cancel':
            pyautogui.confirm('The process has been cancelled.', buttons=['Ok'])
            sys.exit(0)

    # Exception Handling
    except Exception as e:
        pyautogui.confirm('Error has been occurred as: {}.'
                          '\nPlease restart the program.'.format(e))

        sys.exit(0)

# This program is written by:Samish Bhattarai     email: kbhattasamish17@gmail.com
