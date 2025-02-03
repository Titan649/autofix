import subprocess
import win32com.client
import datetime
import os

class AutoFix:
    def __init__(self):
        self.updates_session = win32com.client.Dispatch("Microsoft.Update.Session")
        self.updates_searcher = self.updates_session.CreateUpdateSearcher()

    def list_available_updates(self):
        print("Searching for available updates...")
        search_result = self.updates_searcher.Search("IsInstalled=0")
        updates = search_result.Updates
        update_list = []
        for update in updates:
            update_list.append({
                'Title': update.Title,
                'Description': update.Description,
                'KBArticleIDs': update.KBArticleIDs,
                'MoreInfoUrls': update.MoreInfoUrls,
            })
        return update_list

    def schedule_update(self, update_titles, schedule_time):
        print(f"Scheduling updates: {update_titles} for {schedule_time}")
        update_collection = win32com.client.Dispatch("Microsoft.Update.UpdateColl")
        search_result = self.updates_searcher.Search("IsInstalled=0")
        updates = search_result.Updates

        for update in updates:
            if update.Title in update_titles:
                update_collection.Add(update)

        download_result = self.updates_session.CreateUpdateDownloader().Download(update_collection)
        if download_result.ResultCode == 2:
            print("Updates downloaded successfully. Scheduling installation...")

            task_name = "AutoFix_Update_Installation"
            task_command = f'cmd /c "schtasks /create /tn {task_name} /tr \\"powershell -Command \\"Install-WindowsUpdate -AcceptEula\\"\\" /sc once /st {schedule_time} /f"'
            os.system(task_command)
            print("Updates scheduled successfully.")
        else:
            print("Failed to download updates.")

    def install_updates_now(self, update_titles):
        print("Installing updates...")
        update_collection = win32com.client.Dispatch("Microsoft.Update.UpdateColl")
        search_result = self.updates_searcher.Search("IsInstalled=0")
        updates = search_result.Updates

        for update in updates:
            if update.Title in update_titles:
                update_collection.Add(update)

        installer = self.updates_session.CreateUpdateInstaller()
        installer.Updates = update_collection
        installation_result = installer.Install()
        if installation_result.ResultCode == 2:
            print("Updates installed successfully.")
        else:
            print("Failed to install updates.")

def main():
    autofix = AutoFix()
    print("AutoFix - Advanced Windows Update Control")
    print("1. List Available Updates")
    print("2. Schedule Update Installation")
    print("3. Install Updates Now")
    choice = input("Enter your choice: ")

    if choice == '1':
        updates = autofix.list_available_updates()
        for update in updates:
            print(f"Title: {update['Title']}")
            print(f"Description: {update['Description']}")
            print(f"KB Articles: {update['KBArticleIDs']}")
            print(f"More Info: {update['MoreInfoUrls']}")
            print("-" * 40)

    elif choice == '2':
        update_titles = input("Enter update titles to schedule (comma separated): ").split(',')
        schedule_time = input("Enter schedule time (HH:MM format): ")
        autofix.schedule_update([title.strip() for title in update_titles], schedule_time)

    elif choice == '3':
        update_titles = input("Enter update titles to install (comma separated): ").split(',')
        autofix.install_updates_now([title.strip() for title in update_titles])

    else:
        print("Invalid choice.")

if __name__ == '__main__':
    main()