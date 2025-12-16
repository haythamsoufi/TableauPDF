import os
import logging
import time # Added for task ID simulation
from flask import (
    Flask, render_template, request, jsonify, send_from_directory, url_for,
    flash, redirect
)
from werkzeug.utils import secure_filename
import pandas as pd
from dotenv import load_dotenv
import tableauserverclient as TSC # Import TSC for error handling
from celery import Celery, Task # Import Celery
from celery.result import AsyncResult # Import AsyncResult to check status

# --- Import your adapted backend logic ---
# This now needs to be aware of the Celery app potentially
import backend_logic

# --- App Configuration ---
load_dotenv() # Load environment variables from .env file (for secrets)

# --- Celery Configuration ---
# Option 1: Simple config directly in Flask app
# Replace with your Redis URL if different (e.g., redis://password@hostname:port/0)
CELERY_BROKER_URL = os.environ.get('CELERY_BROKER_URL', 'redis://localhost:6379/0')
CELERY_RESULT_BACKEND = os.environ.get('CELERY_RESULT_BACKEND', 'redis://localhost:6379/0')

# --- Flask App Initialization ---
app = Flask(__name__)
# Secret key needed for flash messages and session management
app.config['SECRET_KEY'] = os.environ.get('FLASK_SECRET_KEY', 'a_very_default_secret_key_change_in_production')
# Configure upload folder (make sure this exists)
UPLOAD_FOLDER = 'uploaded_configs'
ALLOWED_EXTENSIONS = {'xlsx'}
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

# --- Celery App Initialization ---
# Update Celery config within Flask app config
app.config.update(
    CELERY_BROKER_URL=CELERY_BROKER_URL,
    CELERY_RESULT_BACKEND=CELERY_RESULT_BACKEND
)

# Function to create Celery app
def make_celery(app):
    celery = Celery(
        app.import_name,
        backend=app.config['CELERY_RESULT_BACKEND'],
        broker=app.config['CELERY_BROKER_URL']
    )
    # Optional: Configure Celery further (e.g., task routes, rate limits)
    # celery.conf.update(app.config) # Might not be needed with simple config

    # Subclass Task to automatically push Flask app context
    class ContextTask(Task):
        def __call__(self, *args, **kwargs):
            with app.app_context():
                return self.run(*args, **kwargs)

    celery.Task = ContextTask
    return celery

celery_app = make_celery(app) # Create the Celery instance

# --- Import Celery Task ---
# Make sure the task is defined using the @celery_app.task decorator
# We define it in backend_logic.py, so we need to ensure Celery finds it.
# One way is to explicitly import it after celery_app is created.
# Or structure imports carefully. Let's assume backend_logic imports celery_app.
# If using a separate tasks.py: from tasks import run_export_task_celery
# For now, we assume the task definition in backend_logic will use celery_app
# This might require adjusting backend_logic.py slightly to import celery_app
# from backend_logic import run_export_task_celery # This can cause circular imports if backend_logic imports app

# --- Logging Setup ---
logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

# --- Helper Function ---
def allowed_file(filename):
    """Checks if the uploaded file extension is allowed."""
    if '.' in filename:
        extension = filename.rsplit('.', 1)[1].lower()
        return extension in ALLOWED_EXTENSIONS
    return False # No dot means no extension

# --- Routes ---
@app.route('/')
def index():
    """Renders the main page."""
    logger.info("Rendering index page.")
    return render_template('index.html', current_version=backend_logic.CURRENT_VERSION)

@app.route('/upload_excel', methods=['POST'])
def upload_excel():
    """Handles Excel file upload and returns sheet names and first sheet columns."""
    if 'excel_file' not in request.files:
        logger.warning("Excel upload attempt with no file part.")
        return jsonify({'error': 'No file part in the request'}), 400

    file = request.files['excel_file']
    if file.filename == '':
        logger.warning("Excel upload attempt with no selected file.")
        return jsonify({'error': 'No file selected'}), 400

    if file and allowed_file(file.filename):
        try:
            if not os.path.exists(app.config['UPLOAD_FOLDER']):
                 os.makedirs(app.config['UPLOAD_FOLDER'])
                 logger.info(f"Created upload folder: {app.config['UPLOAD_FOLDER']}")

            filename = secure_filename(file.filename)
            # Consider adding timestamp/UUID to prevent overwrites if multiple users upload
            # filename = f"{int(time.time())}_{filename}"
            filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            file.save(filepath)
            logger.info(f"Excel file '{filename}' uploaded successfully to '{filepath}'.")

            sheets = backend_logic.get_excel_sheets(filepath)
            columns = []
            if sheets:
                try:
                    columns = backend_logic.get_excel_columns(filepath, sheets[0])
                except Exception as col_err:
                     logger.warning(f"Could not read columns from first sheet '{sheets[0]}': {col_err}")

            return jsonify({
                'success': True,
                'filename': filename,
                'filepath': filepath, # Send absolute path back
                'sheets': sheets,
                'columns': columns
            })
        except ValueError as ve:
             logger.error(f"Value error processing uploaded Excel file '{file.filename}': {ve}")
             return jsonify({'error': str(ve)}), 400
        except Exception as e:
            logger.error(f"Error processing uploaded Excel file '{file.filename}': {e}", exc_info=True)
            return jsonify({'error': f'Error processing file: {str(e)}'}), 500
    else:
        logger.warning(f"Excel upload attempt with invalid file type: {file.filename}")
        return jsonify({'error': 'Invalid file type. Only .xlsx allowed.'}), 400

# --- Load Views Route (Unchanged conceptually, uses backend_logic) ---
@app.route('/load_views', methods=['POST'])
def load_views():
    """Connects to Tableau and loads view names for a workbook."""
    data = request.json
    logger.info(f"Received request to load views for workbook: {data.get('workbook_name')}")
    try:
        server_url = data.get('server_url')
        token_name = data.get('token_name')
        token_secret = data.get('token_secret') # Secret sent from client
        site_id = data.get('site_id', '')
        workbook_name = data.get('workbook_name')

        if not all([server_url, token_name, token_secret, workbook_name]):
             logger.warning("Load views failed: Missing credentials or workbook name.")
             return jsonify({'error': 'Missing Tableau connection details or workbook name.'}), 400

        logger.warning("Using PAT Secret sent from client for Load Views - replace with secure server-side retrieval!")

        views = backend_logic.get_tableau_views(
            server_url, token_name, token_secret, site_id, workbook_name
        )
        logger.info(f"Successfully loaded {len(views)} views for workbook '{workbook_name}'.")
        return jsonify({'success': True, 'views': views})

    except ValueError as ve:
        logger.error(f"Value error loading Tableau views: {ve}")
        return jsonify({'error': str(ve)}), 400
    except ConnectionError as ce:
         logger.error(f"Connection error loading Tableau views: {ce}")
         return jsonify({'error': str(ce)}), 503
    except Exception as e:
        logger.error(f"Error loading Tableau views: {e}", exc_info=True)
        error_message = f"Failed to connect or load views. Check details and server status."
        return jsonify({'error': error_message}), 500

# --- Get Columns Route (Unchanged conceptually, uses backend_logic) ---
@app.route('/get_columns', methods=['POST'])
def get_columns_for_sheet():
    """Gets column names for a specific sheet in an uploaded Excel file."""
    data = request.json
    filepath = data.get('filepath')
    sheet_name = data.get('sheet_name')
    logger.info(f"Request to get columns for sheet '{sheet_name}' in file '{filepath}'")

    if not filepath or not sheet_name:
        logger.warning("Get columns failed: Missing filepath or sheet name.")
        return jsonify({'error': 'Missing file path or sheet name.'}), 400

    # Security Check: Ensure path is within upload folder
    upload_folder_abs = os.path.abspath(app.config['UPLOAD_FOLDER'])
    try:
        filepath_abs = os.path.abspath(filepath)
    except Exception as path_e:
         logger.error(f"Invalid filepath provided for get_columns: {filepath} - Error: {path_e}")
         return jsonify({'error': 'Invalid file path format.'}), 400

    if not filepath_abs.startswith(upload_folder_abs):
         logger.error(f"Attempt to access file outside upload folder: {filepath}")
         return jsonify({'error': 'Invalid file path.'}), 403 # Forbidden

    if not os.path.exists(filepath_abs):
         logger.error(f"File not found for column retrieval: {filepath_abs}")
         return jsonify({'error': 'File not found. It might have been cleared. Please re-upload.'}), 404

    try:
        columns = backend_logic.get_excel_columns(filepath_abs, sheet_name)
        logger.info(f"Returning {len(columns)} columns for sheet '{sheet_name}'.")
        return jsonify({'success': True, 'columns': columns})
    except ValueError as ve:
        logger.error(f"Value error getting columns for sheet '{sheet_name}': {ve}")
        return jsonify({'error': str(ve)}), 400
    except Exception as e:
        logger.error(f"Error getting columns for sheet '{sheet_name}': {e}", exc_info=True)
        return jsonify({'error': f'Error reading columns: {str(e)}'}), 500

# --- Test Connection Route (Unchanged conceptually, uses backend_logic) ---
@app.route('/test_connection', methods=['POST'])
def test_connection():
    """Tests the Tableau server connection using details from the modal."""
    data = request.json
    logger.info("Received request to test connection.")
    server_url = data.get('server_url')
    token_name = data.get('token_name')
    token_secret = data.get('token_secret') # Secret sent from modal for testing
    site_id = data.get('site_id', '')

    logger.debug(f"Test Connection Data Received: URL='{server_url}', Name='{token_name}', Site='{site_id}', Secret Provided={'Yes' if token_secret else 'No'}")

    if not all([server_url, token_name, token_secret]):
         logger.warning("Test connection failed: Missing credentials.")
         return jsonify({'success': False, 'error': 'Server URL, PAT Name, and PAT Secret are required.'}), 400

    try:
        backend_logic.test_tableau_connection(server_url, token_name, token_secret, site_id)
        logger.info("Test connection successful.")
        return jsonify({'success': True, 'message': 'Connection Successful!'})
    except ValueError as ve:
         logger.warning(f"Test connection validation failed: {ve}")
         return jsonify({'success': False, 'error': str(ve)}), 400
    except ConnectionError as ce:
         logger.error(f"Test connection failed: {ce}", exc_info=False) # Less verbose traceback for connection errors
         return jsonify({'success': False, 'error': str(ce)}), 503
    except Exception as e:
        logger.error(f"Test connection failed unexpectedly: {e}", exc_info=True)
        error_message = f"Connection Failed: An unexpected error occurred. ({type(e).__name__})"
        return jsonify({'success': False, 'error': error_message}), 500

# --- Start Export Route (MODIFIED for Celery) ---
@app.route('/start_export', methods=['POST'])
def start_export():
    """Receives configuration as JSON and triggers the background export task."""
    config_data = request.json
    logger.info("Received request to start export.")

    # --- Server-Side Validation ---
    errors = backend_logic.validate_configuration(config_data)
    if errors:
        logger.warning(f"Export validation failed: {errors}")
        return jsonify({'success': False, 'error': f"Invalid configuration: {'; '.join(errors)}"}), 400

    # --- Securely Get PAT Secret ---
    # *** Replace this with secure retrieval (e.g., environment variable) ***
    pat_secret = config_data.get('token_secret')
    if not pat_secret:
         logger.error("Export failed: PAT Secret missing in request (should be retrieved securely).")
         return jsonify({'success': False, 'error': 'Configuration Error: PAT Secret missing.'}), 400
    else:
         logger.warning("Using PAT Secret sent from client - replace with secure server-side retrieval!")
         # Remove secret from config before passing to task if task gets it separately
         # config_data_for_task = config_data.copy()
         # del config_data_for_task['token_secret']


    # --- Trigger Background Task ---
    try:
        logger.info("Triggering background export task...")
        # Ensure the task function is imported and decorated correctly in backend_logic.py
        # The task name might be 'backend_logic.run_export_task_celery' if defined there
        # Or if you create tasks.py: 'tasks.run_export_task_celery'
        # Using apply_async allows passing arguments
        task = celery_app.send_task(
            'backend_logic.run_export_task_celery', # Name of the task function
            args=[config_data, pat_secret] # Arguments for the task
            # You might add options like queue='export_queue' if using multiple queues
        )

        task_id = task.id
        logger.info(f"Successfully triggered task with ID: {task_id}")
        return jsonify({'success': True, 'message': 'Export process initiated.', 'task_id': task_id})

    except Exception as e:
        logger.error(f"Failed to trigger export task: {e}", exc_info=True)
        return jsonify({'success': False, 'error': 'Failed to start export process on server.'}), 500

# --- Export Status Route (MODIFIED for Celery) ---
@app.route('/export_status/<task_id>')
def export_status(task_id):
     """Checks the status of a background export task using Celery."""
     logger.debug(f"Request for status of task: {task_id}")
     try:
         # Get the task result object from Celery backend
         task_result = AsyncResult(task_id, app=celery_app)
         state = task_result.state

         response_data = {
             'task_id': task_id,
             'status': state,
             'progress': 0,
             'log': [],
             'error': None
         }

         if state == 'PENDING':
             response_data['log'] = ['Task is waiting to start...']
         elif state == 'STARTED':
             response_data['log'] = ['Task has started...']
             # Sometimes progress info is available in STARTED state too
             if isinstance(task_result.info, dict):
                 response_data['progress'] = task_result.info.get('progress', 0)
                 response_data['log'] = task_result.info.get('log', ['Task has started...'])
         elif state == 'PROGRESS':
             if isinstance(task_result.info, dict):
                 response_data['progress'] = task_result.info.get('progress', 0)
                 response_data['log'] = task_result.info.get('log', [])
             else:
                  # Handle cases where info might not be a dict yet
                  response_data['log'] = ['Processing...']
         elif state == 'SUCCESS':
             response_data['progress'] = 100
             if isinstance(task_result.info, dict):
                 response_data['log'] = task_result.info.get('log', ['Task completed successfully.'])
                 # You could potentially return a final result/summary from the task here
                 # response_data['result'] = task_result.info.get('message')
             else:
                  response_data['log'] = ['Task completed successfully.']
         elif state == 'FAILURE':
             response_data['log'] = ['Task failed.']
             # Extract error information if available
             if isinstance(task_result.info, dict):
                  response_data['error'] = task_result.info.get('error', 'Unknown error')
                  response_data['log'] = task_result.info.get('log', ['Task failed.'])
             elif isinstance(task_result.info, Exception):
                  response_data['error'] = str(task_result.info)
             else:
                  response_data['error'] = 'Task failed with unknown error.'
             # Log the traceback on the server for debugging
             logger.error(f"Task {task_id} failed. Traceback: {task_result.traceback}")

         elif state == 'REVOKED':
              response_data['status'] = 'STOPPED' # Use a clearer status for frontend
              response_data['log'] = ['Task was stopped.']
              if isinstance(task_result.info, dict):
                   response_data['progress'] = task_result.info.get('progress', 0) # Show last known progress

         return jsonify(response_data)

     except Exception as e:
         logger.error(f"Error retrieving status for task {task_id}: {e}", exc_info=True)
         # Return an error status if checking fails
         return jsonify({'task_id': task_id, 'status': 'ERROR', 'error': 'Could not retrieve task status.'}), 500


if __name__ == '__main__':
    # Make sure the upload folder exists
    if not os.path.exists(UPLOAD_FOLDER):
        try:
            os.makedirs(UPLOAD_FOLDER)
            logger.info(f"Created upload folder: {UPLOAD_FOLDER}")
        except OSError as e:
             logger.error(f"Could not create upload folder {UPLOAD_FOLDER}: {e}", exc_info=True)
             # Decide if the app should exit if it can't create the folder

    # Set debug=True for development ONLY.
    # Important: For Celery, running Flask in debug mode with the reloader can sometimes cause issues
    # with task discovery or duplicate task execution. It's often better to run Flask with debug=False
    # when testing Celery integration, or set use_reloader=False.
    app.run(debug=False, host='127.0.0.1', port=5001) # Changed debug to False for Celery testing
