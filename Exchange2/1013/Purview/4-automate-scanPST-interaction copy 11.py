from pywinauto import Application
import time
import os
import shutil
import logging
from datetime import datetime
import re

def setup_logging(folder_path):
    """Setup logging configuration"""
    log_dir = os.path.join(folder_path, "logs")
    os.makedirs(log_dir, exist_ok=True)
    log_file = os.path.join(log_dir, f"pst_repair_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log")
    
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler(log_file),
            logging.StreamHandler()
        ]
    )
    return logging.getLogger('PST_Repair')

def extract_numbers_from_text(text):
    """Extract numbers from text using regex"""
    numbers = re.findall(r'\d+', text)
    return [int(num) for num in numbers]

def get_scan_progress(window):
    """Try to extract progress information from the scanning window"""
    try:
        # Look for progress bar or text containing numbers
        for control in window.children():
            text = control.window_text()
            if '%' in text or 'items' in text.lower():
                return text
    except:
        pass
    return None

def wait_for_window(app, title, timeout=30, check_interval=0.5, logger=None):
    """Enhanced window waiting with progress logging"""
    start_time = time.time()
    last_log_time = start_time
    while time.time() - start_time < timeout:
        try:
            window = app.window(title=title)
            if window.exists():
                # Try to get progress information
                progress = get_scan_progress(window)
                if progress and logger and time.time() - last_log_time >= 2:  # Log every 2 seconds
                    logger.info(f"Progress: {progress}")
                    last_log_time = time.time()
                return window
        except Exception as e:
            if logger:
                logger.debug(f"Waiting for window '{title}': {str(e)}")
        time.sleep(check_interval)
    raise TimeoutError(f"Window '{title}' not found within {timeout} seconds")

def capture_repair_stats(window):
    """Capture repair statistics from the dialog"""
    stats = {}
    try:
        # Look for text containing statistics
        for control in window.children():
            text = control.window_text().lower()
            if 'items' in text or 'found' in text or 'repaired' in text:
                stats['repair_details'] = text
                numbers = extract_numbers_from_text(text)
                if numbers:
                    stats['items_processed'] = numbers[0]
    except Exception as e:
        stats['error'] = f"Failed to capture stats: {str(e)}"
    return stats

def calculate_timeout(file_size_mb):
    """Calculate appropriate timeout based on file size
    For large files, we need much more generous timeouts:
    - Base 30 minutes for files up to 1GB
    - Additional 30 minutes per GB
    - Maximum 24 hours for extremely large files
    """
    base_timeout = 1800  # 30 minutes base
    
    # Add 30 minutes (1800 seconds) for each GB
    additional_timeout = (file_size_mb / 1024) * 1800
    
    # Minimum 30 minutes, maximum 24 hours (86400 seconds)
    total_timeout = min(max(base_timeout + additional_timeout, base_timeout), 86400)
    
    return total_timeout

def is_process_responding(window):
    """Check if the SCANPST process is still responding"""
    try:
        # Try to get window text - this will fail if window is not responding
        window.window_text()
        return True
    except:
        return False

def repair_pst(scanpst_path, pst_path, logger):
    app = None
    try:
        # Calculate file size and appropriate timeout
        file_size_mb = os.path.getsize(pst_path) / (1024*1024)
        max_scan_wait = calculate_timeout(file_size_mb)
        
        logger.info(f"Launching SCANPST for: {pst_path}")
        logger.info(f"File size: {file_size_mb:.2f} MB")
        logger.info(f"Timeout set to: {max_scan_wait/3600:.1f} hours")
        
        app = Application().start(f'"{scanpst_path}"')
        
        # Wait for and handle main window
        main_window = wait_for_window(app, 'Microsoft Outlook Inbox Repair Tool', timeout=30, logger=logger)
        logger.info("Main window found")
        
        # Set PST path
        edit_field = main_window.Edit
        edit_field.set_text(pst_path)
        logger.info("PST path entered")
        
        # Click Start
        main_window['&Start'].click()
        logger.info("Started scanning...")
        
        scan_start_time = time.time()
        last_progress_report = 0
        last_activity_check = time.time()
        last_window_check = time.time()
        
        while time.time() - scan_start_time < max_scan_wait:
            current_time = time.time()
            
            try:
                # Process activity checks (from new version)
                if current_time - last_activity_check >= 300:
                    process_active = False
                    for window in app.windows():
                        if is_process_responding(window):
                            process_active = True
                            break
                    
                    if process_active:
                        logger.info("SCANPST process is still responding")
                        last_activity_check = current_time
                
                # Progress monitoring (from new version)
                for window in app.windows():
                    try:
                        if current_time - last_progress_report >= 300:
                            progress = get_scan_progress(window)
                            if progress:
                                logger.info(f"Current progress: {progress}")
                                last_progress_report = current_time
                    except:
                        continue
                
                # Look for repair dialog
                repair_dialog = wait_for_window(app, 'Microsoft Outlook Inbox Repair Tool', timeout=5, logger=logger)
                repair_button = repair_dialog.child_window(title="&Repair", class_name="Button")
                
                if repair_button.exists() and repair_button.is_enabled():
                    logger.info("Scan completed, repair button found")
                    
                    # Handle backup checkbox
                    try:
                        backup_checkbox = repair_dialog.child_window(
                            title="&Make backup of scanned file before repairing",
                            class_name="Button"
                        )
                        if backup_checkbox.exists():
                            backup_checkbox.check()
                            logger.info("Backup checkbox checked")
                    except Exception as e:
                        logger.warning(f"Backup checkbox interaction error: {str(e)}")
                    
                    logger.info("Starting repair...")
                    repair_button.click()
                    
                    # Wait for completion dialog with original handling
                    repair_wait_start = time.time()
                    while time.time() - repair_wait_start < max_scan_wait:
                        try:
                            for window in app.windows():
                                if window.window_text() == "Microsoft Outlook Inbox Repair Tool":
                                    # Look for static text directly
                                    for control in window.children():
                                        if control.window_text() == "Repair complete":
                                            # Find OK button
                                            for button in window.children():
                                                if button.window_text() == "OK":
                                                    logger.info("\nFound completion dialog, clicking OK")
                                                    button.click()
                                                    time.sleep(1)
                                                    return True, None
                                    
                            # Log repair progress periodically
                            if time.time() - last_progress_report >= 300:
                                logger.info("Repair is still in progress...")
                                last_progress_report = time.time()
                                
                        except Exception as e:
                            logger.debug(f"Error in completion check: {str(e)}")
                        time.sleep(0.5)
                    
                    logger.error("Repair completion dialog not found within timeout")
                    return False, None
                    
            except Exception as e:
                logger.debug(f"Error in repair loop: {str(e)}")
            time.sleep(0.5)
        
        logger.error(f"Operation timed out after {max_scan_wait/3600:.1f} hours")
        return False, None
        
    except Exception as e:
        logger.error(f"Error repairing {pst_path}: {str(e)}")
        return False, None
    finally:
        if app:
            try:
                logger.info("Checking window status before cleanup...")
                active_repair = False
                for window in app.windows():
                    try:
                        if is_process_responding(window) and 'scanpst' in window.window_text().lower():
                            active_repair = True
                            logger.warning("Active repair process detected - not forcing close")
                            break
                    except:
                        continue
                
                if not active_repair:
                    logger.info("Cleaning up windows...")
                    for window in app.windows():
                        try:
                            window.close()
                        except:
                            pass
            except Exception as e:
                logger.debug(f"Error in cleanup: {str(e)}")

def batch_repair_psts(folder_path):
    logger = setup_logging(folder_path)
    scanpst_path = r"C:\Program Files (x86)\Microsoft Office\Office16\SCANPST.EXE"
    
    if not os.path.exists(folder_path):
        logger.error(f"Error: Folder path {folder_path} does not exist!")
        return []
    
    logger.info("Starting PST repair process...")
    logger.info(f"Looking for PST files in: {folder_path}")
    
    results = []
    for file in os.listdir(folder_path):
        if file.lower().endswith('.pst'):
            pst_path = os.path.join(folder_path, file)
            logger.info(f"\nProcessing: {pst_path}")
            file_size_mb = os.path.getsize(pst_path) / (1024*1024)
            logger.info(f"File size: {file_size_mb:.2f} MB")
            
            try:
                success, repair_stats = repair_pst(scanpst_path, pst_path, logger)
                result = {
                    'file': file,
                    'status': 'Success' if success else 'Failed',
                    'size_mb': f"{file_size_mb:.2f}",
                    'timestamp': datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    'repair_stats': repair_stats if repair_stats else {}
                }
                results.append(result)
                
                # Handle cleanup
                if success:
                    manage_backup_files(pst_path, logger)
                
            except Exception as e:
                logger.error(f"Error processing {file}: {str(e)}")
                results.append({
                    'file': file,
                    'status': 'Failed',
                    'size_mb': f"{file_size_mb:.2f}",
                    'timestamp': datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    'error': str(e)
                })
            
            logger.info("Waiting before processing next file...")
            time.sleep(5)
    
    return results

def manage_backup_files(pst_path, logger):
    """Enhanced backup file management with logging"""
    try:
        # Handle BAK file
        original_bak = f"{pst_path[:-4]}.bak"
        if os.path.exists(original_bak):
            next_number = get_next_bak_number(pst_path)
            new_bak_name = f"{pst_path[:-4]}_{next_number}.bak"
            shutil.move(original_bak, new_bak_name)
            logger.info(f"Renamed backup file to: {os.path.basename(new_bak_name)}")

        # Handle LOG file
        log_file = f"{pst_path[:-4]}.log"
        if os.path.exists(log_file):
            if wait_for_file_release(log_file):
                log_dir = os.path.join(os.path.dirname(pst_path), "logs")
                os.makedirs(log_dir, exist_ok=True)

                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                base_log_name = os.path.basename(pst_path)[:-4]
                new_log_name = f"{base_log_name}_{timestamp}.log"
                new_log_path = os.path.join(log_dir, new_log_name)

                shutil.move(log_file, new_log_path)
                logger.info(f"Moved log file to: logs/{new_log_name}")
            else:
                logger.warning("Could not access log file - it may still be in use")

        return True
    except Exception as e:
        logger.error(f"Error managing backup files: {str(e)}")
        return False

if __name__ == "__main__":
    folder_path = r"C:\ExchangeArchives"
    results = batch_repair_psts(folder_path)
    
    # Print results
    print("\nRepair Results:")
    print("{:<50} {:<10} {:<15} {:<20} {:<30}".format(
        "File", "Status", "Size (MB)", "Timestamp", "Repair Statistics"))
    print("-" * 125)
    for result in results:
        stats_str = str(result.get('repair_stats', ''))[:30]
        print("{:<50} {:<10} {:<15} {:<20} {:<30}".format(
            result['file'], 
            result['status'], 
            result['size_mb'],
            result['timestamp'],
            stats_str
        ))
        
    # Save results to CSV
    csv_path = os.path.join(folder_path, f"pst_repair_results_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv")
    with open(csv_path, 'w', newline='') as csvfile:
        fieldnames = ['file', 'status', 'size_mb', 'timestamp', 'repair_stats']
        writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(results)
    print(f"\nResults saved to: {csv_path}")