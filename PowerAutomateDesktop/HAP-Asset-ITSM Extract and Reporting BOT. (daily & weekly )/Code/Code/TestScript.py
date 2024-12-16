import win32com.client
def Move():
    def move_emails(folder_name, target_folder_name, static_subject_part):
        try:
            outlook = win32com.client.Dispatch("Outlook.Application")
            namespace = outlook.GetNamespace("MAPI")

            # Get the Inbox folder (top-level)
            inbox = namespace.GetDefaultFolder(win32com.client.constants.olFolderInbox)

            # Get the source folder (subfolder of Inbox)
            source_folder = inbox.Folders[folder_name]
            # Get the target folder (subfolder of Inbox)
            target_folder = inbox.Folders[target_folder_name]

            messages = source_folder.Items
            #messages.Sort("[ReceivedTime]", True)  # Sort by received time, descending

            i = len(messages)  
            print(i)# Initialize the counter to the number of messages
            while i > 0:
                message = messages[i]
                # Check if the static part of the subject is in the email subject
                if static_subject_part.lower() in message.Subject.lower():
                    # Move the email to the target folder
                    message.Move(target_folder)
                    print(f"Moved email with subject: {message.Subject}")
                i -= 1  # Decrement the counter to move to the previous message

            print("Emails moved successfully.")

        except Exception as e:
            print(f"An error occurred: {e}")

    # Usage example:
    folder_name = "RGT_New"  # Source folder
    target_folder_name = "RGT_Processed"  # Target folder
    static_subject_part = "RGT"  # Replace with the static part of the subject
    move_emails(folder_name, target_folder_name, static_subject_part)


def Move1():
    def move_emails(folder_name, target_folder_name, static_subject_part):

        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")

        # Get the Inbox folder (top-level)
        inbox = namespace.GetDefaultFolder(win32com.client.constants.olFolderInbox)

        # Get the source folder (subfolder of Inbox)
        source_folder = inbox.Folders[folder_name]
        # Get the target folder (subfolder of Inbox)
        target_folder = inbox.Folders[target_folder_name]

        messages = source_folder.Items
        #messages.Sort("[ReceivedTime]", True)  # Sort by received time, descending

        i = len(messages)  
        print(i)# Initialize the counter to the number of messages
        while i > 0:
            message = messages[i]
            # Check if the static part of the subject is in the email subject
            if static_subject_part.lower() in message.Subject.lower():
                # Move the email to the target folder
                message.Move(target_folder)
                print(f"Moved email with subject: {message.Subject}")
            i -= 1  # Decrement the counter to move to the previous message

        print("Emails moved successfully.")

    # Usage example:
    folder_name = "RGT_New"  # Source folder
    target_folder_name = "RGT_Processed"  # Target folder
    static_subject_part = "RGT"  # Replace with the static part of the subject
    move_emails(folder_name, target_folder_name, static_subject_part)
