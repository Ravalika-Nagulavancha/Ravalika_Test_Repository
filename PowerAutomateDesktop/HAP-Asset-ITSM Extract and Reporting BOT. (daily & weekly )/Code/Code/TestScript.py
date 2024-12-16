import win32com.client
url: =["https://github.com"]
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

# Usage example:
folder_name = "RGT_New"  # Source folder
target_folder_name = "RGT_Processed"  # Target folder
static_subject_part = "RGT"  # Replace with the static part of the subject
move_emails(folder_name, target_folder_name, static_subject_part)

# Usage example:
folder_name = "RGT_New"  # Source folder
target_folder_name = "RGT_Processed"  # Target folder
static_subject_part = "RGT"  # Replace with the static part of the subject
move_emails(folder_name, target_folder_name, static_subject_part)
