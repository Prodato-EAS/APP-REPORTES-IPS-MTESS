from gevent import monkey
monkey.patch_all()

from flask import Flask, render_template, jsonify, request, send_file, session, redirect, url_for
from flask_session import Session
from sharepoint_manager import SharePointManager
import pandas as pd
import threading
import os
import io
import msal
import time
from flask_socketio import SocketIO, emit, join_room, leave_room
from dotenv import load_dotenv
from datetime import datetime

# Load Env
load_dotenv()

app = Flask(__name__)
app.config['SECRET_KEY'] = os.urandom(24)
app.config['SESSION_TYPE'] = 'filesystem'
MAINTENANCE_MODE = False  

# --- MIDDLEWARE & ERROR HANDLERS ---
@app.before_request
def check_maintenance():
    if MAINTENANCE_MODE:
        if request.endpoint in ['static', 'maintenance_page']:
            return
        return redirect(url_for('maintenance_page'))

@app.errorhandler(404)
def page_not_found(e):
    return render_template('404.html'), 404

@app.route('/maintenance')
def maintenance_page():
    if not MAINTENANCE_MODE:
        return redirect(url_for('index'))
    return render_template('maintenance.html')
Session(app)

socketio = SocketIO(app, cors_allowed_origins="*", async_mode='gevent', manage_session=False)

# Auth Config
CLIENT_ID = os.getenv("CLIENT_ID")
CLIENT_SECRET = os.getenv("CLIENT_SECRET")
TENANT_ID = os.getenv("TENANT_ID")
AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
REDIRECT_PATH = "/getAToken"
ENDPOINT = 'https://graph.microsoft.com/v1.0/me'
SCOPE = ["User.Read"]

# Presence State
connected_users = {} # {sid: {name: str, photo: str, email: str}}

# Singleton Manager
manager = SharePointManager()
import gevent.event
monitor_event = gevent.event.Event()

def _build_auth_url(authority=None, scopes=None, state=None):
    return _build_msal_app(authority=authority).get_authorization_request_url(
        scopes or [],
        state=state or os.urandom(16).hex(),
        prompt="select_account", # Force account picker
        redirect_uri=url_for("authorized", _external=True, _scheme="https"))

def _build_msal_app(cache=None, authority=None):
    return msal.ConfidentialClientApplication(
        CLIENT_ID, authority=authority or AUTHORITY,
        client_credential=CLIENT_SECRET, token_cache=cache)

def background_monitor():
    """Checks for updates (Event-Driven) and pushes to clients."""
    last_emitted_version = 0
    print(" * Background Monitor Started (Event-Driven)")
    last_emitted_versions = {"IPS": 0, "MTESS": 0}
    while True:
        try:
            # Wait for explicit signal OR timeout (external check)
            monitor_event.wait(timeout=10) 
            monitor_event.clear()

            # Sync with version.json
            current_times, updated_sources = manager.check_version()
            
            # Emit updates for specific sources that changed
            for source in ["IPS", "MTESS"]:
                # If newer than what we last broadcasted
                if current_times[source] > (last_emitted_versions[source] + 0.001):
                    # Valid update found
                    df_current = manager.dfs.get(source)
                    new_date_str = get_formatted_date(df_current)
                    
                    socketio.emit('server_update', {
                        "source": source,
                        "version": current_times[source],
                        "last_updated": new_date_str
                    })
                    last_emitted_versions[source] = current_times[source]
                
        except Exception as e:
            print(f"Monitor Error: {e}")
            socketio.sleep(5)

# Start Background Task
socketio.start_background_task(background_monitor)

# --- Auth Routes ---

@app.route("/login")
def login():
    session["state"] = os.urandom(16).hex()
    auth_url = _build_auth_url(scopes=SCOPE, state=session["state"])
    return render_template("login.html", auth_url=auth_url)

@app.route(REDIRECT_PATH)
def authorized():
    if request.args.get('state') != session.get("state"):
        return redirect(url_for("index"))
    
    if "error" in request.args:
        return render_template("login.html", error=request.args.get("error_description"))

    if request.args.get('code'):
        cache = msal.SerializableTokenCache()
        result = _build_msal_app(cache=cache).acquire_token_by_authorization_code(
            request.args['code'],
            scopes=SCOPE,
            redirect_uri=url_for("authorized", _external=True, _scheme="https"))
        
        if "error" in result:
             return render_template("login.html", error=result.get("error_description"))
             
        session["user"] = result.get("id_token_claims")
        
        # Whitelist Check
        email = session["user"].get("preferred_username", "").lower()
        
        # Admin Logic
        ADMIN_EMAILS = [
            "jonathan.arce@prodato.com.py",
            "valeria.penayo@prodato.com.py",
            "andrea.ruiz@prodato.com.py"
        ]
        is_admin = (email in ADMIN_EMAILS)
        session["is_admin"] = is_admin

        # Fetch Access List from SharePoint
        allowed_users = []
        try:
            whitelist_items = manager.get_whitelist() # Returns [{'id': '1', 'email': '...'}, ...]
            allowed_users = [u['email'] for u in whitelist_items]
        except Exception as e:
            print(f"Error checking whitelist: {e}")
        
        # Master Override for Admin
        if not is_admin and email not in allowed_users:
             session.clear()
             return render_template("login.html", error="Acceso denegado: no tienes permiso para acceder a este sitio.")

        # Fetch Photo (Base64)
        try:
             token = result['access_token']
             import requests
             photo_res = requests.get(
                 'https://graph.microsoft.com/v1.0/me/photo/$value',
                 headers={'Authorization': 'Bearer ' + token},
                 stream=True
             )
             if photo_res.status_code == 200:
                 import base64
                 session["photo"] = "data:image/jpeg;base64," + base64.b64encode(photo_res.content).decode('utf-8')
                 session["photo_bytes"] = base64.b64encode(photo_res.content).decode('utf-8') # For JSON socket
             else:
                 # Fallback Initials
                 name = session["user"].get("name", "User")
                 initials = "".join([n[0] for n in name.split()[:2]]).upper()
                 # SVG Placeholder - minimal
                 svg = f'<svg xmlns="http://www.w3.org/2000/svg" width="32" height="32" viewBox="0 0 32 32"><rect width="32" height="32" rx="16" fill="#dee2e6"/><text x="50%" y="50%" dy=".35em" text-anchor="middle" font-family="Arial" font-size="14" fill="#6c757d">{initials}</text></svg>'
                 import base64
                 session["photo"] = "data:image/svg+xml;base64," + base64.b64encode(svg.encode('utf-8')).decode('utf-8')

        except Exception as e:
            print(f"Photo Fetch Error: {e}")
            session["photo"] = "" # Allow empty

    return redirect(url_for("index"))

@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))

@app.route('/')
def index():
    if not session.get("user"):
        return redirect(url_for("login"))
    return render_template('index.html', user=session["user"], photo=session.get("photo"), is_admin=session.get("is_admin", False))

# --- Admin API ---

@app.route('/api/admin/users', methods=['GET'])
def get_users():
    if not session.get("user") or not session.get("is_admin"): return jsonify({"error": "Unauthorized"}), 403
    try:
        users = manager.get_whitelist()
        return jsonify({"result": "success", "users": users})
    except Exception as e:
        return jsonify({"result": "error", "message": str(e)}), 500

@app.route('/api/admin/users', methods=['POST'])
def add_user():
    if not session.get("user") or not session.get("is_admin"): return jsonify({"error": "Unauthorized"}), 403
    try:
        data = request.json
        email = data.get("email", "").strip().lower()
        
        if not email: return jsonify({"result": "error", "message": "Email requerido"}), 400
        if not email.endswith("@prodato.com.py"): 
            return jsonify({"result": "error", "message": "Solo se permiten correos @prodato.com.py"}), 400
        
        # Check duplicate
        current = manager.get_whitelist()
        if any(u['email'] == email for u in current):
            return jsonify({"result": "error", "message": "El usuario ya existe"}), 400

        success, msg = manager.add_to_whitelist(email)
        if success: return jsonify({"result": "success", "message": msg})
        else: return jsonify({"result": "error", "message": msg}), 500

    except Exception as e:
        return jsonify({"result": "error", "message": str(e)}), 500

@app.route('/api/admin/users/<item_id>', methods=['DELETE'])
def delete_user(item_id):
    if not session.get("user") or not session.get("is_admin"): return jsonify({"error": "Unauthorized"}), 403
    try:
        success, msg = manager.remove_from_whitelist(item_id)
        if success: return jsonify({"result": "success", "message": msg})
        else: return jsonify({"result": "error", "message": msg}), 500

    except Exception as e:
        return jsonify({"result": "error", "message": str(e)}), 500

# --- Presence (SocketIO) ---

@socketio.on('connect')
def handle_connect():
    if not session.get("user"):
        return False # Reject
    
    user_info = {
        "name": session["user"].get("name"),
        "email": session["user"].get("preferred_username"),
        "photo": session.get("photo")
    }
    
    is_new_person = True
    for sid, u in connected_users.items():
        if u['email'] == user_info['email']:
            is_new_person = False
            break
            
    connected_users[request.sid] = user_info
    
    if is_new_person:
        emit('user_joined', user_info, broadcast=True, include_self=False)
    
    # Send full list to SELF always
    all_users = list(connected_users.values())
    emit('presence_list', all_users)

@socketio.on('disconnect')
def handle_disconnect():
    if request.sid in connected_users:
        user = connected_users.pop(request.sid)
        is_still_here = False
        for u in connected_users.values():
            if u['email'] == user['email']:
                is_still_here = True
                break
        
        if not is_still_here:
             emit('user_left', user, broadcast=True) 
        all_users = list(connected_users.values())
        emit('presence_update', all_users, broadcast=True)

# --- Standard API ---

@app.route('/api/update', methods=['POST'])
def update_status():
    if not session.get("user"): return jsonify({"error": "Unauthorized"}), 401
    try:
        data = request.json
        ids = data.get('ids', [])
        status = data.get('status', 'REVISADO')
        source = data.get('source', 'IPS') # Default to IPS
        
        if not ids:
            return jsonify({"result": "error", "message": "No IDs provided"}), 400

        modifier_name = session.get("user", {}).get("name", "Usuario")
        modifier_email = session.get("user", {}).get("preferred_username", "")
        
        manager.update_status_by_ids(ids, status, modifier=modifier_name, source=source)
        
        manager.patch_sharepoint_background(ids, status=status, modifier_name=modifier_name, source=source)
        
        new_date_str = get_formatted_date(manager.dfs.get(source))
        
        monitor_event.set()

        return jsonify({"result": "success", "count": len(ids), "last_updated": new_date_str})
    except Exception as e:
        print(f"Server Error in /api/update: {e}")
        return jsonify({"result": "error", "message": f"Server Error: {str(e)}"}), 500


def get_formatted_date(df):
    date_str = "Desconocido"
    if not df.empty and "Modified" in df.columns:
        valid_dates = df["Modified"].dropna()
        if not valid_dates.empty:
            max_date = valid_dates.max()
            try:
                if max_date.tzinfo is None:
                    max_date = max_date.tz_localize('UTC')
                max_date_py = max_date.tz_convert('America/Asuncion')
                
                days = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado", "Domingo"]
                months = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"]
                
                day_name = days[max_date_py.dayofweek]
                month_name = months[max_date_py.month-1]
                time_str = max_date_py.strftime("%H:%M")
                
                date_str = f"{day_name}, {max_date_py.day} de {month_name} de {max_date_py.year} a las {time_str} horas"
            except Exception as e:
                date_str = str(max_date)
    return date_str

@app.route('/api/data')
def get_data():
    if not session.get("user"): return jsonify({"error": "Unauthorized"}), 401
    
    source = request.args.get('source', 'IPS') # Default 'IPS'
    
    manager.check_version()
    
    df = manager.dfs.get(source, pd.DataFrame())

    refresh = request.args.get('refresh', 'false')
    if refresh == 'true' or df.empty:
        df = manager.fetch_data(source)
    
    date_str = get_formatted_date(df)

    df_json = df.copy()
    if "Modified" in df_json.columns:
        df_json["Modified"] = df_json["Modified"].astype(str) 
    
    df_json = df_json.fillna("")
    return jsonify({
        "data": df_json.to_dict(orient='records'),
        "last_updated": date_str
    })

# ... update status ...

@app.route('/api/export')
def export_excel():
    if not session.get("user"): return redirect(url_for("login"))
    company = request.args.get('company', 'Todas')
    source = request.args.get('source', 'IPS')

    df_export = manager.dfs.get(source, pd.DataFrame()).copy()
    if company != 'Todas':
        if "RUC" in df_export.columns:
            df_export = df_export[df_export["RUC"].astype(str) == company]
        else:
            df_export = df_export[df_export["Empresa"].astype(str) == company]

    # Filter Inconsistencies 
    bad_states = ['NO_COINCIDE', 'SIN_REGISTRO_IPS', 'SIN_REGISTRO_MTESS']
    mask_incons = (
        (df_export["AUD_ENTRADA"].isin(bad_states)) | 
        (df_export["AUD_ESTADO"].isin(bad_states)) | 
        (df_export["AUD_SALIDA"].isin(bad_states))
    )
    df_incons = df_export[mask_incons]

    cols_to_drop = [c for c in ["#", "Modified", "ID", "REVISADO", "ModificadoPor"] if c in df_export.columns]
    df_export = df_export.drop(columns=cols_to_drop, errors='ignore')
    df_incons = df_incons.drop(columns=cols_to_drop, errors='ignore')

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_export.to_excel(writer, sheet_name="Data_General", index=False)
        if not df_incons.empty:
            df_incons.to_excel(writer, sheet_name="Data_Inconsistencias", index=False)
            
    output.seek(0)
    
    company_name = "General"
    if company != 'Todas':
        if not df_export.empty and "Empresa" in df_export.columns:
            company_name = str(df_export.iloc[0]["Empresa"]).strip()
        else:
            company_name = str(company)

    safe_name = "".join([c for c in company_name if c.isalnum() or c in (' ', '-', '_')]).strip()
    date_str = datetime.now().strftime("%d-%m-%Y")
    filename = f"Auditoria_{safe_name}_{date_str}.xlsx"

    return send_file(output, download_name=filename, as_attachment=True)

@app.route('/api/version')
def get_version():
    return jsonify({"version": manager.check_version()})

@app.route('/api/export/pdf')
def export_pdf():
    if not session.get("user"): return "Unauthorized", 401
    
    company_ruc = request.args.get('company', '')
    source = request.args.get('source', 'IPS')
    
    if not company_ruc or company_ruc == 'Todas':
        return "PDF export requires a specific company selection", 400
    
    try:
        from datetime import datetime, timezone, timedelta
        from pdf_generator import generate_pdf_reportlab
        
        df = manager.dfs.get(source, pd.DataFrame())
        if df.empty:
            df = manager.fetch_data(source)
        
        company_data = df[df['RUC'] == company_ruc].copy()
        if company_data.empty:
            return "No data found for this company", 404
        
        company_name = company_data.iloc[0]['Empresa'] if not company_data.empty else "Sin Nombre"
        
        def is_inconsistent(row):
            bad_states = ['NO_COINCIDE', 'SIN_REGISTRO_IPS', 'SIN_REGISTRO_MTESS']
            return (row['AUD_ENTRADA'] in bad_states or 
                    row['AUD_ESTADO'] in bad_states or 
                    row['AUD_SALIDA'] in bad_states)
        
        inconsistencies = company_data[company_data.apply(is_inconsistent, axis=1)]
        
        # Generate PDF using ReportLab
        pdf_buffer = generate_pdf_reportlab(company_data, inconsistencies, company_ruc, company_name)
        
        # Prepare filename
        py_tz = timezone(timedelta(hours=-3))
        now = datetime.now(py_tz)
        safe_name = "".join([c for c in company_name if c.isalnum() or c in (' ', '-', '_')]).strip()
        date_str = now.strftime("%d-%m-%Y")
        filename = f"Reporte_{safe_name}_{date_str}.pdf"
        
        # Return PDF
        return send_file(pdf_buffer, as_attachment=False, download_name=filename, mimetype='application/pdf')
    
    except Exception as e:
        print(f"PDF Error: {e}")
        import traceback
        traceback.print_exc()
        return f"Error generating PDF: {str(e)}", 500


if __name__ == '__main__':
    def initial_fetch():
         manager.fetch_data("IPS")
         manager.fetch_data("MTESS")
    threading.Thread(target=initial_fetch).start()
    port = int(os.environ.get('PORT', 8000))
    print(f" * Running on http://0.0.0.0:{port}")
    socketio.run(app, host='0.0.0.0', port=port, debug=False)