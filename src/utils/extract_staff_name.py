import subprocess

def extract_staff_name(auuid_digit):
    try:
        # Execute the command to get user information
        command = f'net user /domain "{auuid_digit}"'
        result = subprocess.run(command, capture_output=True, text=True, shell=True)

        # Check if the command was successful
        if result.returncode == 0:
            # Parse the output to find the staff name
            output_lines = result.stdout.splitlines()
            for line in output_lines:
                if "Full Name" in line:
                    # Extract the full name from the line
                    full_name = line.split(":", 1)[1].strip()
                    return full_name
        else:
            print(f"Error executing command for AUUID {auuid_digit}: {result.stderr}")
            return None
    except Exception as e:
        print(f"An error occurred: {e}")
        return None