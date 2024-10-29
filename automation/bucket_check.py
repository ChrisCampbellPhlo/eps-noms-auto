def list_bucket_files(bucket_name):
    """
    Lists all files in the specified GCP bucket.
    """
    try:
        storage_client = storage.Client()
        bucket = storage_client.bucket(bucket_name)
        blobs = bucket.list_blobs()
        
        files = [blob.name for blob in blobs]
        print(f"Found {len(files)} bucket contents {bucket_name}")
        return files
    except Exception as e:
        print(f"Error checking bucket contents: {e}")
        return []

def check_file_exists(bucket_name, blob_name):
    """
    Existing file check.
    """
    try:
        storage_client = storage.Client()
        bucket = storage_client.bucket(bucket_name)
        blob = bucket.blob(blob_name)
        
        exists = blob.exists()
        if exists:
            print(f"File {blob_name} already exists {bucket_name}")
        return exists
    except Exception as e:
        print(f"Error checking file existence: {e}")
        return False

def get_latest_processed_date(bucket_name):
    """
    Get most recent date
    """
    try:
        files = list_bucket_files(bucket_name)
        processed_files = [f for f in files if f.startswith('processed_eps_nom_report+')]
        
        if not processed_files:
            print("No processed files found in bucket")
            return None
        
        # Extract dates from filenames + convert to datetime
        dates = []
        for filename in processed_files:
            try:
                date_str = filename.split('+')[1][:6]  # Extract YYMMDD
                date = datetime.strptime(date_str, '%y%m%d')
                dates.append(date)
            except (IndexError, ValueError):
                continue
        
        if dates:
            latest_date = max(dates)
            print(f"Latest processed file date: {latest_date.strftime('%Y-%m-%d')}")
            return latest_date
        return None
    except Exception as e:
        print(f"Error getting latest processed date: {e}")
        return None

### modify main folder to with bucket check functions

def main():
    """
    Main function with added GCP bucket checking
    """
    # Set up auth
    authentication()
    
    # Config
    base_url = "https://digital.nhs.uk/services/electronic-prescription-service/statistics"
    base_filename = "eps_nom_report+"
    gcp_bucket_name = "phlo-sandpit-raw-data-lake/sources/reference-data/nhs-eps-noms"

    # Check latest file in GCP
    latest_processed_date = get_latest_processed_date(gcp_bucket_name)
    
    # Get the latest report date
    report_date = get_latest_report_date()
    
    # Skip if file exists
    if latest_processed_date and report_date <= latest_processed_date:
        print("Already have the latest file processed. Skipping download.")
        return
    
    source_filename = generate_filename(base_filename, report_date)
    excel_url = f"{base_url}/{source_filename}"
    
    local_excel_filename = source_filename
    modified_excel_filename = f"modified_{source_filename}"
    local_csv_filename = f"{source_filename[:-5]}.csv"
    gcp_csv_blob_name = f"processed_{source_filename[:-5]}.csv"

    # Check if file exists
    if check_file_exists(gcp_bucket_name, gcp_csv_blob_name):
        print(f"File {gcp_csv_blob_name} already exists in GCP. Skipping processing.")
        return
    

    if download_excel(excel_url, local_excel_filename):
        modified_workbook = modify_excel(local_excel_filename)
        if modified_workbook is not None:
            if save_excel(modified_workbook, modified_excel_filename):
                if excel_to_csv(modified_workbook, local_csv_filename):
                    if upload_to_gcp(gcp_bucket_name, local_csv_filename, gcp_csv_blob_name):
                        print("Process successfully complete")
                    else:
                        print("Failed to upload CSV to GCP. Exiting Process.")
                else:
                    print("Failed to convert xlsx to CSV. Exiting Process.")
            else:
                print("Failed to save modified Excel file. Exiting Process.")
        else:
            print("Failed to modify Excel file. Exiting Process.")
    else:
        print("Failed to download Excel file. Exiting.")

    # Clean up
    try:
        os.remove(local_excel_filename)
        os.remove(modified_excel_filename)
        os.remove(local_csv_filename)
        print("Local files removed")
    except Exception as e:
        print(f"Error removing local files: {e}")
