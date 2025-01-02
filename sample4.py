import os
import subprocess
import streamlit as st


def compress_folder(folder_path, output_zip_path):
    try:
        subprocess.run([
            "powershell", 
            "-Command", 
            f"Compress-Archive -Path '{folder_path}' -DestinationPath '{output_zip_path}'"
        ], check=True)
        return True, f"Folder successfully compressed to {output_zip_path}"
    except subprocess.CalledProcessError as e:
        return False, f"Error compressing folder: {e}"


def send_email_with_attachment(zip_path, recipient_email):
    try:
        powershell_script = f"""
        $Outlook = New-Object -ComObject Outlook.Application
        $Mail = $Outlook.CreateItem(0)
        $Mail.Subject = "zip file: $(Get-Date -Format 'dd-MMM-yyyy')"
        $body ="<br>Dear Team,</br>"
        $body +="<br>Please find the attached file as of $(Get-Date -Format 'dd-MMM-yyyy')</br>"
        $body +="<br>Best Regards</br>"
        $body +="<br>DS</br>"
        $Mail.HTMLBody=$body
        $Mail.To = "{recipient_email}"
        $Mail.CC = "chakraborty.barnil@dxc.com;"
        $Mail.BCC = ""
        $Mail.Attachments.Add("{zip_path}")
        $Mail.Send()
        """
        subprocess.run(["powershell", "-Command", powershell_script], check=True)
        return True, "Email sent successfully!"
    except subprocess.CalledProcessError as e:
        return False, f"Error sending email: {e}"


st.title("Folder to ZIP and Email Sender")

folder_path = st.text_input("Enter the folder path you want to compress:")


if st.button("Compress Folder"):
    if folder_path:
        folder_name = os.path.basename(folder_path.rstrip("/\\"))
        output_zip_path = os.path.join("C:\\", f"{folder_name}.zip")
        success, result_message = compress_folder(folder_path, output_zip_path)
        st.write(result_message)

        if success:
            # Store the zip path in session state
            st.session_state.zip_path = output_zip_path
    else:
        st.write("Please enter a valid folder path.")

if 'zip_path' in st.session_state:
    st.write(f"ZIP file created at {st.session_state.zip_path}.")
    
    recipient_email = st.text_input("Enter recipient's email address:")

    if st.button("Send Email"):
        if recipient_email:
            success, email_result = send_email_with_attachment(st.session_state.zip_path, recipient_email)
            st.write(email_result)
        else:
            st.write("Please enter a valid email address.")
