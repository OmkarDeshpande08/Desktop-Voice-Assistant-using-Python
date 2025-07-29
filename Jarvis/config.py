"""
Configuration file for Jarvis Assistant
"""

class config:
    # Wolfram Alpha App ID - Replace with your actual app ID
    wolframalpha_id = "YOUR_WOLFRAM_ALPHA_APP_ID"
    
    # Default paths for different operating systems
    import platform
    
    if platform.system() == "Windows":
        # Windows paths
        office_path = "C:\\Program Files\\Microsoft Office\\root\\Office16\\"
        resume_path = "C:/Users/Downloads/resume.pdf"
        screenshot_path = "D://JARVIS//JARVIS_2.0//"
    else:
        # Linux/Unix paths
        office_path = "/usr/bin/"  # LibreOffice path
        resume_path = "~/Downloads/resume.pdf"
        screenshot_path = "./screenshots/"
    
    # Voice settings
    voice_rate = 200
    voice_volume = 0.9
    
    # Camera settings
    camera_index = 0
    
    # Search paths for folders (cross-platform)
    if platform.system() == "Windows":
        folder_search_paths = ["C:\\", "D:\\", "E:\\", "F:\\"]
    else:
        folder_search_paths = ["/home", "/usr", "/opt", "/var"]
    
    # Application commands for different platforms
    if platform.system() == "Windows":
        apps = {
            "notepad": "notepad.exe",
            "calculator": "calc.exe",
            "cmd": "cmd.exe"
        }
    else:
        apps = {
            "notepad": "gedit",  # or "nano", "vim"
            "calculator": "gnome-calculator",  # or "kcalc"
            "cmd": "gnome-terminal"  # or "xterm"
        }