import psutil


def kill_processes(process_name: str):
    """
    Kill all processes with the given name.
    Returns True if at least one process was killed, False otherwise.
    """
    pids: list[int] = get_matching_processes(process_name)
    if len(pids) == 0:
        print(f"No processes found with name '{process_name}'.")
        return False

    process: psutil.Process = None
    for pid in pids:
        try:
            process = psutil.Process(pid)
            # process.terminate()  # Send SIGTERM (graceful termination)
            # print(f"Terminated process '{process_name}' with PID: {pid}")
            # # Optional: Wait briefly to ensure termination
            # process.wait(timeout=3)
            process.kill()
        except psutil.NoSuchProcess:
            print(f"Process with PID {pid} no longer exists.")
        except psutil.AccessDenied:
            print(f"Access denied to terminate process with PID {pid}. Try running as administrator/root.")
        except psutil.TimeoutExpired:
            print(f"Process with PID {pid} did not terminate in time. Forcing kill...")
            process.kill()  # Send SIGKILL (forceful termination)
            print(f"Forced kill of process with PID {pid}.")
        except Exception as e:
            print(f"Error terminating process with PID {pid}: {e}")

    return True


def get_matching_processes(process_name: str) -> list[int]:
    """
    Check if a process with the given name is running.
    Returns True if found, False otherwise.
    """
    pids: list[int] = []
    for process in psutil.process_iter(['name']):

        try:
            # Compare process name (case-insensitive)
            if process.info['name'].lower().startswith(process_name.lower()):
                pids.append(process.pid)

        except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
            # Skip processes that can't be accessed or no longer exist
            continue

    return pids
