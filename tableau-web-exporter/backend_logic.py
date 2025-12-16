import logging
import os
import pandas as pd
import time # Ensure time is imported
import tableauserverclient as TSC
from tableauserverclient.server.endpoint.exceptions import NotSignedInError, ServerResponseError # Import specific errors


# --- Constants (Can be moved to config) ---
CURRENT_VERSION = "v1.11-Web-LS-v2" # Example version
logger = logging.getLogger(__name__) # Use logger configured in app.py


try:
    from app import celery_app # Assumes celery_app is defined in app.py
except ImportError:
    # Fallback if running script directly or celery_app not found
    logger.warning("Could not import celery_app from app.py. Celery task decorator will be skipped.")
    celery_app = None # Define as None so the decorator check doesn't fail

# --- Custom Exception for Export Errors (Optional) ---
class ExportError(Exception):
    """Custom exception for export-related failures."""
    pass

# --- Excel Functions ---
def get_excel_sheets(filepath):
    """Gets sheet names from an Excel file."""
    try:
        logger.info(f"Reading sheets from: {filepath}")
        if not os.path.exists(filepath):
             raise FileNotFoundError(f"Excel file not found at path: {filepath}")
        xls = pd.ExcelFile(filepath)
        sheets = xls.sheet_names
        logger.info(f"Found sheets: {sheets}")
        return sheets
    except FileNotFoundError:
        logger.error(f"Excel file not found at: {filepath}")
        raise
    except Exception as e:
        logger.error(f"Error reading sheets from {filepath}: {e}", exc_info=True)
        raise ValueError(f"Could not read sheets from the provided file: {e}")

def get_excel_columns(filepath, sheet_name):
    """Gets column names from a specific sheet."""
    try:
        logger.info(f"Reading columns from sheet '{sheet_name}' in {filepath}")
        if not os.path.exists(filepath):
             raise FileNotFoundError(f"Excel file not found at path: {filepath}")
        df_header = pd.read_excel(filepath, sheet_name=sheet_name, nrows=0)
        columns = list(df_header.columns)
        logger.info(f"Found columns: {columns}")
        return columns
    except FileNotFoundError:
         logger.error(f"Excel file not found at path: {filepath}")
         raise
    except Exception as e:
        logger.error(f"Error reading columns from sheet '{sheet_name}': {e}", exc_info=True)
        if isinstance(e, KeyError) or "Worksheet named" in str(e) or "No sheet named" in str(e):
             raise ValueError(f"Sheet named '{sheet_name}' not found in the Excel file.")
        raise ValueError(f"Could not read columns from sheet '{sheet_name}': {e}")


# --- Tableau Functions ---
def get_tableau_views(server_url, token_name, token_secret, site_id, workbook_name_to_find):
    """Connects to Tableau and gets view names for a specific workbook."""
    logger.info(f"Connecting to Tableau: {server_url}, Site: '{site_id or 'Default'}', Workbook: '{workbook_name_to_find}'")
    server = None
    if not server_url: raise ValueError("Server URL cannot be empty.")
    if not workbook_name_to_find: raise ValueError("Workbook Name cannot be empty.")
    if not token_name or not token_secret: raise ValueError("PAT Name and Secret are required.")

    if not server_url.startswith(('http://', 'https://')):
        server_url = 'https://' + server_url
        logger.info(f"Prepended 'https://' to server URL.")

    signed_in = False
    try:
        auth = TSC.PersonalAccessTokenAuth(token_name, token_secret, site_id=site_id)
        server = TSC.Server(server_url, use_server_version=True)
        server.add_http_options({'timeout': 120})
        server.auth.sign_in_with_personal_access_token(auth)
        signed_in = True
        logger.info("Tableau sign-in successful.")

        target_workbook = find_workbook(server, workbook_name_to_find)
        if not target_workbook:
            raise ValueError(f"Workbook '{workbook_name_to_find}' not found on site '{site_id or 'Default'}'. Check name and site ID.")

        logger.info(f"Found workbook '{target_workbook.name}'. Populating views...")
        server.workbooks.populate_views(target_workbook)
        view_names = sorted([view.name for view in target_workbook.views if not getattr(view, 'is_hidden', False)])
        logger.info(f"Found {len(view_names)} non-hidden views.")
        return view_names

    except TSC.ServerResponseError as sre:
         logger.error(f"Tableau Server Response Error during view loading: {sre.code} - {sre.summary} - {sre.detail}", exc_info=False)
         if str(sre.code) == "401":
              raise ConnectionError("Authentication failed. Check PAT Name/Secret and Site ID.")
         elif str(sre.code) == "404":
              raise ValueError(f"Server or Site not found (404). Check URL: {server_url} and Site ID: '{site_id or 'Default'}'.")
         else:
              raise ConnectionError(f"Tableau server error ({sre.code}): {sre.summary}") from sre
    except ConnectionError as ce:
         logger.error(f"Network or connection error loading Tableau views: {ce}", exc_info=True)
         raise ConnectionError(f"Network error connecting to Tableau: {ce}") from ce
    except Exception as e:
        logger.error(f"Unexpected error loading Tableau views: {e}", exc_info=True)
        raise ConnectionError(f"Failed to connect or load views: {e}") from e
    finally:
        if server and signed_in:
            try:
                server.auth.sign_out()
                logger.info("Tableau sign out successful.")
            except Exception as sign_out_e:
                logger.error(f"Error during Tableau sign out: {sign_out_e}")

def test_tableau_connection(server_url, token_name, token_secret, site_id):
    """Attempts to sign in and sign out of Tableau Server."""
    logger.info(f"Attempting test connection to {server_url} (Site: '{site_id or 'Default'}')")
    server = None
    signed_in = False

    if not server_url:
        raise ValueError("Server URL cannot be empty.")
    if not server_url.startswith(('http://', 'https://')):
        server_url = 'https://' + server_url
        logger.info(f"Prepended 'https://' to server URL: {server_url}")

    if not token_name or not token_secret:
         raise ValueError("PAT Name and Secret cannot be empty.")

    try:
        auth = TSC.PersonalAccessTokenAuth(token_name, token_secret, site_id=site_id)
        server = TSC.Server(server_url, use_server_version=True)
        server.add_http_options({'timeout': 30})
        logger.debug("Attempting sign in...")
        server.auth.sign_in_with_personal_access_token(auth)
        signed_in = True
        logger.info("Test connection: Sign in successful.")

        try:
            logger.debug("Attempting sign out...")
            server.auth.sign_out()
            logger.info("Test connection: Sign out successful.")
        except Exception as sign_out_e:
             logger.warning(f"Test connection: Sign out failed, but sign in was successful. Error: {sign_out_e}")

        return True

    except TSC.ServerResponseError as sre:
         logger.error(f"Test connection failed during sign in attempt: {sre.code} - {sre.summary}", exc_info=False)
         if str(sre.code) == "401":
              raise ConnectionError("Authentication failed. Check PAT Name/Secret and Site ID.")
         elif str(sre.code) == "404":
              raise ConnectionError(f"Server or Site not found (404). Check URL: {server_url} and Site ID: '{site_id or 'Default'}'.")
         else:
              raise ConnectionError(f"Tableau server error ({sre.code}): {sre.summary}") from sre
    except ConnectionError as ce:
         logger.error(f"Network error during test connection: {ce}", exc_info=True)
         raise ConnectionError(f"Network error during test connection: {ce}") from ce
    except Exception as e:
        logger.error(f"Test connection failed during sign in attempt: {e}", exc_info=True)
        raise ConnectionError(f"Tableau sign-in failed: {e}") from e

def validate_configuration(config_data):
    """Performs server-side validation of the configuration dictionary."""
    errors = []
    logger.debug("Validating configuration server-side.")

    required_server = ['server_url', 'token_name', 'workbook_name']
    for key in required_server:
        if not config_data.get(key):
            errors.append(f"Missing required server setting: {key}.")

    required_export = ['export_mode', 'export_format']
    for key in required_export:
        if not config_data.get(key):
             errors.append(f"Missing required export setting: {key}.")
        elif key == 'export_format' and config_data[key] not in ['PDF', 'PNG']:
             errors.append(f"Invalid export_format: {config_data[key]}. Must be 'PDF' or 'PNG'.")

    if config_data.get('export_mode') == 'automate':
        excel_filepath = config_data.get('excel_filepath')
        sheet_name = config_data.get('sheet_name')

        if not excel_filepath:
            errors.append("Missing Excel file path for automate mode.")
        elif not os.path.exists(excel_filepath):
             logger.warning(f"Excel file path provided but not found: {excel_filepath}")
             errors.append(f"Uploaded Excel file not found. Please re-upload.")
        elif not sheet_name:
             errors.append("Missing Sheet Name for automate mode.")
        else:
             try:
                 xls = pd.ExcelFile(excel_filepath)
                 if sheet_name not in xls.sheet_names:
                     errors.append(f"Sheet '{sheet_name}' not found in the uploaded Excel file.")
                 else:
                     df_cols = list(pd.read_excel(excel_filepath, sheet_name=sheet_name, nrows=0).columns)
                     if not df_cols:
                          errors.append(f"Sheet '{sheet_name}' appears to have no columns or header row.")
                     else:
                         # Validate fields used in config against actual columns
                         fields_to_check = [
                             config_data.get('tableau_filter_field'),
                             config_data.get('file_naming_option') if config_data.get('file_naming_option') != 'By view' else None,
                             config_data.get('organize_by_1') if config_data.get('organize_by_1') != 'None' else None,
                             config_data.get('organize_by_2') if config_data.get('organize_by_2') != 'None' else None,
                         ]
                         for f in config_data.get('filters', []): fields_to_check.append(f.get('field'))
                         for c in config_data.get('conditions', []): fields_to_check.append(c.get('field'))

                         for field in filter(None, fields_to_check):
                             if field and field not in df_cols:
                                 errors.append(f"Configured column '{field}' not found in sheet '{sheet_name}'.")

             except Exception as e:
                 logger.error(f"Error validating Excel file structure ({excel_filepath}, sheet: {sheet_name}): {e}")
                 errors.append("Could not validate Excel file structure. Please ensure it's a valid .xlsx file and the sheet exists.")

    # Check filter/condition value presence based on type
    for i, cond in enumerate(config_data.get('conditions', [])):
        cond_type = cond.get('type')
        cond_val = cond.get('value')
        if cond_type not in ['Is Blank', 'Is Not Blank'] and (cond_val is None or str(cond_val).strip() == ""):
            errors.append(f"Condition #{i+1}: Value is missing or empty for type '{cond_type}'.")

    return errors


# --- Helper Functions for Export Task ---
def find_workbook(server, workbook_name):
    """Finds a workbook by name, case-insensitive fallback."""
    logger.debug(f"Helper: Finding workbook '{workbook_name}'...")
    req_option = TSC.RequestOptions(pagesize=1)
    req_option.filter.add(TSC.Filter(TSC.RequestOptions.Field.Name, TSC.RequestOptions.Operator.Equals, workbook_name))
    all_matching_workbooks, _ = server.workbooks.get(req_option)
    if all_matching_workbooks:
        logger.info(f"Found workbook by exact match: {all_matching_workbooks[0].name}")
        return all_matching_workbooks[0]
    else:
        logger.warning(f"Workbook '{workbook_name}' not found (exact match). Trying case-insensitive...")
        req_option_all = TSC.RequestOptions(pagesize=1000) # Consider limiting?
        all_workbooks, _ = server.workbooks.get(req_option_all)
        target_workbook = next((wb for wb in all_workbooks if wb.name.lower() == workbook_name.lower()), None)
        if target_workbook:
             logger.info(f"Found workbook by case-insensitive match: {target_workbook.name}")
        else:
             logger.error(f"Workbook '{workbook_name}' not found after case-insensitive search.")
        return target_workbook

def apply_filters(df, filters_config):
    """Applies Excel row filters based on UI config."""
    if not filters_config:
        logger.info("No Excel filters defined.")
        return df

    logger.info(f"Applying {len(filters_config)} Excel list filters.")
    filtered_df = df.copy()
    try:
        for i, filt in enumerate(filters_config):
            field = filt.get('field')
            values_str = filt.get('values_str', '')
            if not field:
                logger.warning(f"Skipping Filter #{i+1}: Missing 'field'.")
                continue
            if not values_str:
                logger.warning(f"Skipping Filter #{i+1} for field '{field}': Missing 'values_str'.")
                continue

            selected_values = {str(v).strip() for v in values_str.split(',') if v.strip()}
            if not selected_values:
                 logger.warning(f"Skipping Filter #{i+1} for field '{field}': No valid values provided in '{values_str}'.")
                 continue

            logger.debug(f"Applying Filter #{i+1}: Field='{field}', Values='{selected_values}'")
            if field not in filtered_df.columns:
                logger.warning(f"Skipping Filter #{i+1}: Field '{field}' not found in DataFrame columns.")
                continue

            original_len = len(filtered_df)
            mask = filtered_df[field].astype(str).isin(selected_values)
            filtered_df = filtered_df[mask]
            rows_removed = original_len - len(filtered_df)
            logger.debug(f"Filter #{i+1} applied. Rows remaining: {len(filtered_df)} (-{rows_removed})")

            if filtered_df.empty:
                logger.info("DataFrame empty after applying filter, stopping further filtering.")
                break
    except Exception as e:
        logger.error(f"Error applying Excel filters: {e}", exc_info=True)
        raise ValueError(f"Failed to apply Excel filters: {e}")

    return filtered_df

def check_condition(item_row, condition_config):
    """Checks if a single condition is met for a given item row."""
    cond_col = condition_config.get('field')
    cond_type = condition_config.get('type', 'Equals')
    cond_val_str_ui = condition_config.get('value', '')

    if not cond_col:
        logger.warning(f"Skipping condition: Missing 'field'. Config: {condition_config}")
        return False

    actual_value_raw = item_row.get(cond_col, None)

    is_actual_blank = False
    actual_value_str_stripped = ""
    if actual_value_raw is None or pd.isna(actual_value_raw):
         is_actual_blank = True
    else:
         actual_value_str_stripped = str(actual_value_raw).strip()
         if actual_value_str_stripped == "":
              is_actual_blank = True

    condition_met = False
    try:
        if cond_type == 'Is Blank':
            condition_met = is_actual_blank
        elif cond_type == 'Is Not Blank':
            condition_met = not is_actual_blank
        elif cond_type in ['Equals', 'Not Equals']:
            if is_actual_blank:
                 match = cond_val_str_ui.strip() == ""
            else:
                 try:
                     if actual_value_str_stripped == "" or cond_val_str_ui.strip() == "":
                          match = actual_value_str_stripped.lower() == cond_val_str_ui.strip().lower()
                     else:
                          actual_num = float(actual_value_raw)
                          cond_num = float(cond_val_str_ui)
                          match = actual_num == cond_num
                 except (ValueError, TypeError):
                     match = actual_value_str_stripped.lower() == cond_val_str_ui.strip().lower()
            condition_met = match if cond_type == 'Equals' else not match
        elif cond_type in ['Greater Than', 'Less Than']:
             if is_actual_blank or cond_val_str_ui.strip() == "":
                  condition_met = False
             else:
                  try:
                     actual_num = pd.to_numeric(actual_value_raw, errors='coerce')
                     cond_num = pd.to_numeric(cond_val_str_ui, errors='coerce')
                     if pd.notna(actual_num) and pd.notna(cond_num):
                         if cond_type == 'Greater Than': condition_met = actual_num > cond_num
                         elif cond_type == 'Less Than': condition_met = actual_num < cond_num
                     else:
                          condition_met = False
                  except (ValueError, TypeError):
                       condition_met = False
        else:
             logger.warning(f"Unsupported condition type: {cond_type}")
             condition_met = False

    except Exception as e:
        logger.error(f"Error evaluating condition (Col: {cond_col}, Type: {cond_type}, Val: {cond_val_str_ui}, Actual: {actual_value_raw}): {e}", exc_info=True)
        condition_met = False

    logger.debug(f"Condition Check: Col='{cond_col}', Type='{cond_type}', UIVal='{cond_val_str_ui}', Actual='{actual_value_str_stripped}', IsBlank={is_actual_blank}, Met={condition_met}")
    return condition_met

def determine_views_for_item(all_views_in_workbook, item_row, config_data):
    """Determines which views to export for a specific item row."""
    logger.debug("Helper: Determining views for item...")
    global_exclusions = set(config_data.get('excluded_views', []))
    conditional_exclusions = set()
    conditions = config_data.get('conditions', [])

    for i, condition in enumerate(conditions):
        try:
            # Ensure condition has necessary fields before checking
            if condition.get('field') and condition.get('type') and 'excluded_views_str' in condition:
                if check_condition(item_row, condition):
                    excluded_views_str = condition.get('excluded_views_str', '')
                    excluded_for_cond = {v.strip() for v in excluded_views_str.split(',') if v.strip()}
                    if excluded_for_cond:
                         logger.debug(f"Condition #{i+1} met. Excluding views: {excluded_for_cond}")
                         conditional_exclusions.update(excluded_for_cond)
            else:
                 logger.warning(f"Skipping invalid condition structure at index {i}: {condition}")
        except Exception as e:
             logger.error(f"Error processing condition #{i+1} for item: {e}", exc_info=True)

    final_exclusions = global_exclusions.union(conditional_exclusions)
    views_to_export = [view for view in all_views_in_workbook if view.name not in final_exclusions]
    logger.debug(f"Final views to export for item ({len(views_to_export)}): {[v.name for v in views_to_export]}")
    return views_to_export


def determine_parameters(item_row, config_data):
    """Determines the parameter overrides for a specific item row."""
    logger.debug("Helper: Determining parameters for item...")
    parameters_config = config_data.get('parameters', [])
    item_parameters = {}

    for i, param_config in enumerate(parameters_config):
        param_name = param_config.get('name', '').strip()
        param_value_source = param_config.get('value', '') # Keep original value source

        if not param_name:
            logger.warning(f"Skipping Parameter #{i+1}: Name is empty.")
            continue
        # Value can be empty, so don't skip based on that

        final_param_value = param_value_source # Default to static

        # Check if the source exists as a column in the item_row
        if param_value_source in item_row.index:
            cell_value = item_row.get(param_value_source) # Use .get() for safety
            if cell_value is None or pd.isna(cell_value):
                 final_param_value = "" # Use empty string for blank Excel cells
                 logger.debug(f"Parameter '{param_name}': Using empty string (source column '{param_value_source}' was blank/NA).")
            else:
                 final_param_value = str(cell_value)
                 logger.debug(f"Parameter '{param_name}': Using value from column '{param_value_source}' -> '{final_param_value}'.")
        else:
            # Treat as a static value if not a column name
             logger.debug(f"Parameter '{param_name}': Using static value '{param_value_source}'.")

        item_parameters[param_name] = final_param_value

    logger.debug(f"Final parameters for item: {item_parameters}")
    return item_parameters

def sanitize_filename(name):
    """Removes potentially problematic characters for filenames/folders."""
    name = str(name).strip()
    # Allow letters, numbers, space, underscore, hyphen, period
    sanitized = "".join(c if c.isalnum() or c in (' ', '_', '-', '.') else '_' for c in name)
    # Replace multiple consecutive underscores/spaces with a single one
    sanitized = "_".join(filter(None, sanitized.split('_')))
    sanitized = " ".join(filter(None, sanitized.split(' ')))
    max_len = 100
    if len(sanitized) > max_len:
        logger.warning(f"Sanitized name '{sanitized}' exceeded max length {max_len}, truncating.")
        sanitized = sanitized[:max_len].strip('_ ')
    if not sanitized:
        sanitized = "invalid_name"
        logger.warning(f"Name '{name}' became empty after sanitization, using fallback.")
    return sanitized

def determine_output_path(base_path, item_row, config_data):
    """Determines the full output path including subfolders for an item."""
    logger.debug("Helper: Determining output path for item...")
    org1_col = config_data.get('organize_by_1', 'None')
    org2_col = config_data.get('organize_by_2', 'None')

    path_parts = [base_path]

    if org1_col != 'None':
         org1_val_raw = item_row.get(org1_col)
         if org1_val_raw is None and org1_col in item_row.index:
              org1_val = sanitize_filename("NA_Folder")
              logger.warning(f"Organize By 1 column '{org1_col}' is NA/blank, using '{org1_val}'.")
         elif org1_val_raw is not None:
              org1_val = sanitize_filename(org1_val_raw)
              if org1_val: path_parts.append(org1_val)
         else:
              logger.warning(f"Organize By 1 column '{org1_col}' not found in item row.")

    if org2_col != 'None':
         org2_val_raw = item_row.get(org2_col)
         if org2_val_raw is None and org2_col in item_row.index:
              org2_val = sanitize_filename("NA_Folder")
              logger.warning(f"Organize By 2 column '{org2_col}' is NA/blank, using '{org2_val}'.")
         elif org2_val_raw is not None:
              org2_val = sanitize_filename(org2_val_raw)
              if org2_val: path_parts.append(org2_val)
         else:
              logger.warning(f"Organize By 2 column '{org2_col}' not found in item row.")


    final_path = os.path.join(*path_parts)
    # Ensure the path uses the correct OS separator
    final_path = os.path.normpath(final_path)
    logger.debug(f"Determined output path: {final_path}")
    return final_path

def export_single_view(server, view, parameters, output_path, config_data, item_row=None, counter=None):
    """Exports a single view with parameters and naming logic."""
    export_format = config_data.get('export_format', 'PDF')
    file_naming_col = config_data.get('file_naming_option', 'By view')
    numbering_enabled = config_data.get('numbering_enabled', True)

    base_name = view.name # Default
    if item_row is not None and file_naming_col != 'By view':
         if file_naming_col in item_row.index:
             cell_value = item_row.get(file_naming_col)
             if cell_value is None or pd.isna(cell_value) or str(cell_value).strip() == "":
                 logger.warning(f"File naming column '{file_naming_col}' is empty/NA for item, using view name '{view.name}' as fallback.")
             else:
                 base_name = str(cell_value).strip()
         else:
             logger.warning(f"File naming column '{file_naming_col}' not found for item, using view name '{view.name}' as fallback.")

    base_name_clean = sanitize_filename(base_name)
    prefix = f"{counter:02d}_" if numbering_enabled and counter is not None else ""
    file_name_final = f"{prefix}{base_name_clean}.{export_format.lower()}"
    export_file_path = os.path.join(output_path, file_name_final)
    export_file_path = os.path.normpath(export_file_path) # Normalize path

    logger.info(f"Attempting export: View='{view.name}', Path='{export_file_path}', Params={parameters}")

    if export_format == "PDF":
        req_options = TSC.PDFRequestOptions(page_type=TSC.PDFRequestOptions.PageType.Unspecified, maxage=0)
    else: # PNG
        req_options = TSC.ImageRequestOptions(imageresolution=TSC.ImageRequestOptions.Resolution.High, maxage=0)

    tableau_filter_field = config_data.get('tableau_filter_field')
    if item_row is not None and tableau_filter_field and tableau_filter_field not in parameters:
        if tableau_filter_field in item_row.index:
             filter_value_raw = item_row.get(tableau_filter_field)
             filter_value = "" if (filter_value_raw is None or pd.isna(filter_value_raw)) else str(filter_value_raw)
             logger.debug(f"Applying primary Tableau filter: {tableau_filter_field} = '{filter_value}'")
             req_options.vf(tableau_filter_field, filter_value)
        else:
             logger.warning(f"Primary Tableau filter field '{tableau_filter_field}' not found in item row.")

    for name, value in parameters.items():
         logger.debug(f"Applying parameter override: {name} = '{value}'")
         req_options.vf(name, value)

    max_retries = 2
    retries = 0
    success = False
    last_error = None

    while retries <= max_retries and not success:
        try:
            logger.debug(f"Export attempt {retries + 1} for '{view.name}' to {export_file_path}")
            start_export_time = time.time()

            if export_format == "PDF":
                server.views.populate_pdf(view, req_options)
                export_data = view.pdf
            else: # PNG
                server.views.populate_image(view, req_options)
                export_data = view.image

            export_duration = time.time() - start_export_time
            logger.debug(f"View '{view.name}' data populated in {export_duration:.2f}s.")

            # Ensure output directory exists just before writing
            os.makedirs(os.path.dirname(export_file_path), exist_ok=True)

            with open(export_file_path, 'wb') as f:
                f.write(export_data)

            logger.info(f"Successfully exported '{view.name}' as '{file_name_final}' (Attempt {retries + 1}).")
            success = True

        except TSC.ServerResponseError as sre:
             retries += 1
             last_error = sre
             logger.warning(f"Export attempt {retries} failed for view '{view.name}' (Code: {sre.code}): {sre.summary}", exc_info=(retries > max_retries))
             if str(sre.code) == "403":
                  logger.error(f"Permission denied (403) exporting '{view.name}'. Skipping retries.")
                  break
             if retries <= max_retries:
                 wait_time = 3 * retries
                 logger.info(f"Waiting {wait_time}s before retrying view '{view.name}'.")
                 time.sleep(wait_time)
             else:
                 logger.error(f"Max retries reached for '{view.name}'. Skipping this view. Last error: {last_error}")
                 break
        except Exception as e:
            retries += 1
            last_error = e
            logger.warning(f"Export attempt {retries} failed for view '{view.name}': {e}", exc_info=(retries > max_retries))
            if retries <= max_retries:
                wait_time = 3 * retries
                logger.info(f"Waiting {wait_time}s before retrying view '{view.name}'.")
                time.sleep(wait_time)
            else:
                logger.error(f"Max retries reached for '{view.name}'. Skipping this view. Last error: {last_error}")
                break

    if not success:
         logger.error(f"Failed to export '{view.name}' to '{export_file_path}' after multiple retries.")
         raise ExportError(f"Failed to export view '{view.name}' after {max_retries+1} attempts.") from last_error


# --- Celery Task Definition ---
# Add the decorator ONLY if celery_app was imported successfully
task_decorator = celery_app.task(bind=True) if celery_app else lambda f: f

@task_decorator # Apply the decorator (or the dummy lambda if Celery not setup)
def run_export_task_celery(self, config_data, pat_secret): # 'self' is the Celery task instance
    """
    The background task that performs the actual Tableau export.
    """
    is_celery_task = hasattr(self, 'update_state') # Check if running as a real Celery task
    task_id = getattr(self.request, 'id', f'local_run_{int(time.time())}') if is_celery_task else f'direct_run_{int(time.time())}'
    logger.info(f"Task {task_id}: Export task started.")
    log_buffer = []
    last_known_progress = 0

    # --- Helper to update Celery state ---
    def update_progress(progress, message):
        nonlocal log_buffer, last_known_progress
        progress = int(progress)
        last_known_progress = progress
        log_entry = f"[{time.strftime('%H:%M:%S')}] {message}"
        logger.info(f"Task {task_id}: {log_entry} ({progress}%)")
        log_buffer.append(log_entry)
        if len(log_buffer) > 50: log_buffer = log_buffer[-50:]
        if is_celery_task:
            try:
                self.update_state(state='PROGRESS', meta={'progress': progress, 'log': log_buffer})
            except Exception as update_err:
                 logger.error(f"Task {task_id}: Failed to update Celery state: {update_err}")
        else:
             logger.debug("Cannot update Celery state (not running as bound task?).")


    server = None
    start_time = time.time()
    processed_items_count = 0
    final_message = "Process initiated."
    export_summary = {'success': 0, 'failed': 0, 'skipped': 0}

    try:
        # --- 1. Extract Config & Define Output ---
        update_progress(0, "Extracting configuration...")
        server_url = config_data['server_url']
        token_name = config_data['token_name']
        site_id = config_data.get('site_id', '')
        workbook_name = config_data['workbook_name']
        export_mode = config_data['export_mode']
        excel_filepath = config_data.get('excel_filepath')
        sheet_name = config_data.get('sheet_name')
        numbering_enabled = config_data.get('numbering_enabled', True)

        output_folder_base = os.path.abspath(os.path.join(os.getcwd(), "exported_files"))
        logger.info(f"Task {task_id}: Base output directory set to: {output_folder_base}")
        os.makedirs(output_folder_base, exist_ok=True)

        # --- 2. Connect to Tableau ---
        update_progress(5, f"Connecting to Tableau: {server_url}...")
        if not server_url.startswith(('http://', 'https://')): server_url = 'https://' + server_url
        auth = TSC.PersonalAccessTokenAuth(token_name, pat_secret, site_id=site_id)
        server = TSC.Server(server_url, use_server_version=True)
        server.add_http_options({'timeout': 180})
        server.auth.sign_in_with_personal_access_token(auth)
        update_progress(10, "Connected to Tableau.")

        # --- 3. Find Workbook & Views ---
        update_progress(12, f"Finding workbook '{workbook_name}'...")
        target_workbook = find_workbook(server, workbook_name)
        if not target_workbook: raise ValueError(f"Workbook '{workbook_name}' not found.")
        update_progress(15, f"Found workbook '{target_workbook.name}'. Retrieving view list...")
        server.workbooks.populate_views(target_workbook)
        all_views_in_workbook = target_workbook.views
        update_progress(20, f"Found {len(all_views_in_workbook)} views.")

        # --- 4. Process Based on Mode ---
        total_items_to_process = 0

        if export_mode == 'automate':
            update_progress(25, f"Loading Excel data from {excel_filepath}...")
            if not excel_filepath or not os.path.exists(excel_filepath):
                raise FileNotFoundError("Excel file path invalid or file not found on server.")
            df = pd.read_excel(excel_filepath, sheet_name=sheet_name, dtype=str).fillna('')
            update_progress(30, f"Loaded {len(df)} rows from sheet '{sheet_name}'.")

            filters_config = config_data.get('filters', [])
            filtered_df = apply_filters(df, filters_config)
            total_items_to_process = len(filtered_df)
            update_progress(35, f"Processing {total_items_to_process} filtered items.")

            if total_items_to_process == 0:
                 update_progress(100, "No items to process after filtering.")
                 final_message = 'No items to process after filtering.'
            else:
                for index, item_row in filtered_df.iterrows():
                    processed_items_count += 1
                    # Calculate progress: 35% base + 60% spread over items
                    item_progress = 35 + int(60 * (processed_items_count / total_items_to_process))
                    update_progress(item_progress, f"Processing item {processed_items_count}/{total_items_to_process} (Excel Index: {index})...")

                    try:
                        views_to_export = determine_views_for_item(all_views_in_workbook, item_row, config_data)
                        if not views_to_export:
                             update_progress(item_progress, f"Item {processed_items_count}: No views to export after exclusions.")
                             export_summary['skipped'] += 1
                             continue

                        item_parameters = determine_parameters(item_row, config_data)
                        item_output_path = determine_output_path(output_folder_base, item_row, config_data)

                        view_counter = 1
                        item_view_failures = 0
                        for view in views_to_export:
                             try:
                                 export_single_view(server, view, item_parameters, item_output_path, config_data, item_row, view_counter if numbering_enabled else None)
                                 export_summary['success'] += 1
                             except ExportError as single_export_err:
                                  logger.error(f"Item {processed_items_count}: Failed to export view '{getattr(view, 'name', 'N/A')}'. Error: {single_export_err}")
                                  export_summary['failed'] += 1
                                  item_view_failures += 1
                             except Exception as general_err:
                                  logger.error(f"Item {processed_items_count}: Unexpected error exporting view '{getattr(view, 'name', 'N/A')}'. Error: {general_err}", exc_info=True)
                                  export_summary['failed'] += 1
                                  item_view_failures += 1
                             finally:
                                 if numbering_enabled: view_counter += 1
                        if item_view_failures > 0:
                             update_progress(item_progress, f"Item {processed_items_count}: Completed with {item_view_failures} view export failure(s).")

                    except Exception as item_err:
                         logger.error(f"Failed to process item {processed_items_count} (Excel Index: {index}): {item_err}", exc_info=True)
                         update_progress(item_progress, f"Item {processed_items_count}: Failed to process. Error: {item_err}")
                         export_summary['skipped'] += 1

                final_message = f"Processed {processed_items_count} items. Success: {export_summary['success']}, Failed: {export_summary['failed']}, Skipped Items: {export_summary['skipped']}."

        else: # export_mode == 'all_once'
            update_progress(30, "Determining views to export...")
            excluded_views_set = set(config_data.get('excluded_views', []))
            views_to_export = [v for v in all_views_in_workbook if v.name not in excluded_views_set]
            total_items_to_process = len(views_to_export)
            update_progress(40, f"Exporting {total_items_to_process} selected views.")

            if total_items_to_process == 0:
                 update_progress(100, "No views selected for export.")
                 final_message = 'No views selected for export.'
            else:
                view_counter = 1
                for view in views_to_export:
                    processed_items_count += 1
                    # Calculate progress: 40% base + 55% spread over views
                    item_progress = 40 + int(55 * (processed_items_count / total_items_to_process))
                    update_progress(item_progress, f"Exporting view {processed_items_count}/{total_items_to_process}: {view.name}...")
                    try:
                        export_single_view(server, view, {}, output_folder_base, config_data, None, view_counter if numbering_enabled else None)
                        export_summary['success'] += 1
                    except ExportError as single_export_err:
                         logger.error(f"Failed to export view '{getattr(view, 'name', 'N/A')}'. Error: {single_export_err}")
                         export_summary['failed'] += 1
                    except Exception as general_err:
                         logger.error(f"Unexpected error exporting view '{getattr(view, 'name', 'N/A')}'. Error: {general_err}", exc_info=True)
                         export_summary['failed'] += 1
                    finally:
                         if numbering_enabled: view_counter += 1
                final_message = f"Processed {processed_items_count} views. Success: {export_summary['success']}, Failed: {export_summary['failed']}."

        # --- 5. Final Update ---
        end_time = time.time()
        duration = end_time - start_time
        completion_log = f"Export process finished in {duration:.2f} seconds. {final_message}"
        update_progress(100, completion_log)
        logger.info(f"Task {task_id}: {completion_log}")
        # Return final status for Celery result backend
        return {'status': 'Complete', 'message': completion_log, 'log': log_buffer, 'progress': 100}

    except Exception as e:
        logger.error(f"Task {task_id}: Failed! Error: {e}", exc_info=True)
        error_message = f"ERROR: {type(e).__name__} - {e}"
        log_buffer.append(error_message)
        # Use self.update_state when running as a real Celery task
        if is_celery_task:
            try:
                 self.update_state(state='FAILURE', meta={'progress': last_known_progress, 'log': log_buffer, 'error': str(e)})
            except Exception as state_update_err:
                 logger.error(f"Task {task_id}: Also failed to update Celery state for failure: {state_update_err}")
        # It's good practice to raise the exception so Celery knows it failed
        raise e # This will mark the task as FAILED in Celery
    finally:
        # --- Sign Out ---
        if server and getattr(server, 'auth_token', None): # Check if signed in before signing out
            try:
                server.auth.sign_out()
                logger.info(f"Task {task_id}: Signed out from Tableau.")
            except NotSignedInError:
                 logger.info(f"Task {task_id}: Already signed out or sign-in failed.")
            except Exception as sign_out_e:
                logger.error(f"Task {task_id}: Error during Tableau sign out: {sign_out_e}")