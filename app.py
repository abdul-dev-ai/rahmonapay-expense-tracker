from flask import Flask, render_template, request, redirect, url_for, jsonify, send_file, session, flash
import sqlite3
from datetime import datetime, timedelta
import csv
import io
from werkzeug.security import generate_password_hash, check_password_hash
from functools import wraps
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
import os

app = Flask(__name__)
app.secret_key = "rahmonapay-secret-key-2025"
DB_NAME = "expenses.db"
DEFAULT_BUDGET = 0
NEWSLETTER_FILE = "newsletter_subscribers.xlsx"


def init_db():
    """Initialize the database with expenses table"""
    try:
        conn = sqlite3.connect(DB_NAME)
        cur = conn.cursor()
        
        # Users table with newsletter field and phone
        cur.execute("""
            CREATE TABLE IF NOT EXISTS users (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                username TEXT UNIQUE NOT NULL,
                email TEXT UNIQUE NOT NULL,
                phone TEXT,
                password TEXT NOT NULL,
                newsletter_subscribed INTEGER DEFAULT 0,
                sms_alerts INTEGER DEFAULT 0,
                email_alerts INTEGER DEFAULT 1,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        """)
        
        # Expenses table (now with user_id AND budget_type)
        cur.execute("""
            CREATE TABLE IF NOT EXISTS expenses (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                user_id INTEGER NOT NULL,
                budget_type TEXT NOT NULL,
                category TEXT NOT NULL,
                amount REAL NOT NULL,
                date TEXT NOT NULL,
                description TEXT,
                FOREIGN KEY (user_id) REFERENCES users (id)
            )
        """)
        
        # Settings table (now with user_id AND budget_type)
        cur.execute("""
            CREATE TABLE IF NOT EXISTS settings (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                user_id INTEGER NOT NULL,
                budget_type TEXT NOT NULL,
                budget REAL NOT NULL DEFAULT 0,
                UNIQUE(user_id, budget_type),
                FOREIGN KEY (user_id) REFERENCES users (id)
            )
        """)
        
        # Alerts table to track sent alerts
        cur.execute("""
            CREATE TABLE IF NOT EXISTS alerts (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                user_id INTEGER NOT NULL,
                budget_type TEXT NOT NULL,
                alert_type TEXT NOT NULL,
                message TEXT NOT NULL,
                sent_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                FOREIGN KEY (user_id) REFERENCES users (id)
            )
        """)
        
        conn.commit()
        conn.close()
        print("✓ Database initialized successfully")
    except sqlite3.Error as e:
        print(f"✗ Database error: {e}")


def init_newsletter_spreadsheet():
    """Initialize the newsletter subscribers spreadsheet"""
    if not os.path.exists(NEWSLETTER_FILE):
        wb = Workbook()
        sheet = wb.active
        sheet.title = "Newsletter Subscribers"
        
        # Headers
        headers = ['ID', 'Username', 'Email', 'Signup Date', 'Status']
        sheet.append(headers)
        
        # Style headers
        for cell in sheet[1]:
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill(start_color="38bdf8", end_color="38bdf8", fill_type="solid")
            cell.alignment = Alignment(horizontal="center")
        
        # Set column widths
        sheet.column_dimensions['A'].width = 8
        sheet.column_dimensions['B'].width = 20
        sheet.column_dimensions['C'].width = 30
        sheet.column_dimensions['D'].width = 20
        sheet.column_dimensions['E'].width = 12
        
        wb.save(NEWSLETTER_FILE)
        print(f"✓ Newsletter spreadsheet created: {NEWSLETTER_FILE}")


def add_newsletter_subscriber(user_id, username, email):
    """Add a new newsletter subscriber to the Excel spreadsheet"""
    try:
        # Load or create workbook
        if os.path.exists(NEWSLETTER_FILE):
            wb = load_workbook(NEWSLETTER_FILE)
            sheet = wb.active
        else:
            init_newsletter_spreadsheet()
            wb = load_workbook(NEWSLETTER_FILE)
            sheet = wb.active
        
        # Add new subscriber
        signup_date = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        sheet.append([user_id, username, email, signup_date, "Active"])
        
        # Save
        wb.save(NEWSLETTER_FILE)
        print(f"✓ Added newsletter subscriber: {email}")
        return True
    except Exception as e:
        print(f"✗ Error adding newsletter subscriber: {e}")
        return False


def send_email_alert(user_email, username, alert_type, budget_type, total_spent, budget, remaining):
    """Send email alert for budget status"""
    try:
        # In a production environment, you would use an email service like:
        # - SendGrid
        # - Mailgun
        # - AWS SES
        # - SMTP
        
        # For now, we'll just log the alert (you can implement actual email sending later)
        
        if alert_type == "50_percent":
            subject = f"⚠️ RahmonaPay Alert: 50% Budget Reached ({budget_type.title()})"
            message = f"""
Hello {username},

You've reached 50% of your {budget_type} budget!

Budget: GH₵ {budget:.2f}
Spent: GH₵ {total_spent:.2f}
Remaining: GH₵ {remaining:.2f}

Consider monitoring your spending to stay within budget.

Best regards,
RahmonaPay Team
"""
        else:  # over_budget
            subject = f"🚨 RahmonaPay Alert: Budget Exceeded ({budget_type.title()})"
            message = f"""
Hello {username},

WARNING: You have exceeded your {budget_type} budget!

Budget: GH₵ {budget:.2f}
Spent: GH₵ {total_spent:.2f}
Over by: GH₵ {abs(remaining):.2f}

Please review your expenses and adjust your spending.

Best regards,
RahmonaPay Team
"""
        
        print(f"📧 EMAIL ALERT to {user_email}:")
        print(f"   Subject: {subject}")
        print(f"   Message: {message}")
        
        # TODO: Implement actual email sending here
        # Example with SendGrid:
        # from sendgrid import SendGridAPIClient
        # from sendgrid.helpers.mail import Mail
        # message = Mail(
        #     from_email='alerts@rahmonapay.com',
        #     to_emails=user_email,
        #     subject=subject,
        #     plain_text_content=message
        # )
        # sg = SendGridAPIClient(os.environ.get('SENDGRID_API_KEY'))
        # response = sg.send(message)
        
        return True
    except Exception as e:
        print(f"✗ Error sending email alert: {e}")
        return False


def send_sms_alert(phone, username, alert_type, budget_type, total_spent, budget, remaining):
    """Send SMS alert for budget status"""
    try:
        # In a production environment, you would use an SMS service like:
        # - Twilio
        # - Africa's Talking (great for Ghana!)
        # - AWS SNS
        
        # Format phone number (remove spaces, add country code if needed)
        formatted_phone = phone.replace(" ", "").replace("-", "")
        if not formatted_phone.startswith("+"):
            # Assume Ghana country code if not provided
            formatted_phone = "+233" + formatted_phone.lstrip("0")
        
        if alert_type == "50_percent":
            message = f"RahmonaPay Alert: You've reached 50% of your {budget_type} budget. Spent: GH₵{total_spent:.2f} / GH₵{budget:.2f}"
        else:  # over_budget
            message = f"RahmonaPay Alert: Budget exceeded! Spent: GH₵{total_spent:.2f} / GH₵{budget:.2f}. Over by: GH₵{abs(remaining):.2f}"
        
        print(f"📱 SMS ALERT to {formatted_phone}:")
        print(f"   Message: {message}")
        
        # TODO: Implement actual SMS sending here
        # Example with Twilio:
        # from twilio.rest import Client
        # client = Client(account_sid, auth_token)
        # message = client.messages.create(
        #     body=message,
        #     from_='+1234567890',  # Your Twilio number
        #     to=formatted_phone
        # )
        
        # Example with Africa's Talking (recommended for Ghana):
        # import africastalking
        # africastalking.initialize(username, api_key)
        # sms = africastalking.SMS
        # response = sms.send(message, [formatted_phone])
        
        return True
    except Exception as e:
        print(f"✗ Error sending SMS alert: {e}")
        return False


def check_and_send_alerts(user_id, budget_type):
    """Check budget status and send alerts if needed"""
    try:
        conn = sqlite3.connect(DB_NAME)
        cur = conn.cursor()
        
        # Get user info
        cur.execute("""
            SELECT username, email, phone, email_alerts, sms_alerts 
            FROM users WHERE id = ?
        """, (user_id,))
        user = cur.fetchone()
        
        if not user:
            conn.close()
            return
        
        username, email, phone, email_alerts, sms_alerts = user
        
        # Get budget
        cur.execute("""
            SELECT budget FROM settings 
            WHERE user_id = ? AND budget_type = ?
        """, (user_id, budget_type))
        budget_result = cur.fetchone()
        budget = budget_result[0] if budget_result else 0
        
        if budget <= 0:
            conn.close()
            return
        
        # Get expenses
        cur.execute("""
            SELECT SUM(amount) FROM expenses 
            WHERE user_id = ? AND budget_type = ?
        """, (user_id, budget_type))
        total_spent = cur.fetchone()[0] or 0
        
        remaining = budget - total_spent
        percentage_spent = (total_spent / budget) * 100
        
        # Check if we should send alerts
        alert_type = None
        
        if total_spent > budget:
            alert_type = "over_budget"
        elif percentage_spent >= 50:
            alert_type = "50_percent"
        
        if alert_type:
            # Check if we already sent this alert recently (within last 24 hours)
            cur.execute("""
                SELECT id FROM alerts 
                WHERE user_id = ? 
                AND budget_type = ? 
                AND alert_type = ? 
                AND sent_at > datetime('now', '-1 day')
            """, (user_id, budget_type, alert_type))
            
            recent_alert = cur.fetchone()
            
            if not recent_alert:
                # Send alerts based on user preferences
                if email_alerts:
                    send_email_alert(email, username, alert_type, budget_type, 
                                   total_spent, budget, remaining)
                
                if sms_alerts and phone:
                    send_sms_alert(phone, username, alert_type, budget_type, 
                                 total_spent, budget, remaining)
                
                # Log the alert
                message = f"Alert sent: {alert_type} for {budget_type} budget"
                cur.execute("""
                    INSERT INTO alerts (user_id, budget_type, alert_type, message) 
                    VALUES (?, ?, ?, ?)
                """, (user_id, budget_type, alert_type, message))
                
                conn.commit()
                print(f"✓ Alert sent to user {username} for {budget_type} budget")
        
        conn.close()
        
    except Exception as e:
        print(f"✗ Error checking/sending alerts: {e}")


def get_expenses():
    """Retrieve all expenses from database with ID, ordered by date (newest first)"""
    if 'user_id' not in session or 'budget_type' not in session:
        return []
    
    try:
        conn = sqlite3.connect(DB_NAME)
        cur = conn.cursor()
        cur.execute("SELECT category, amount, date, id, description FROM expenses WHERE user_id = ? AND budget_type = ? ORDER BY date DESC, id DESC", 
                   (session['user_id'], session['budget_type']))
        rows = cur.fetchall()
        conn.close()
        return rows
    except sqlite3.Error as e:
        print(f"Error fetching expenses: {e}")
        return []


def get_budget():
    """Get the current budget from settings"""
    if 'user_id' not in session or 'budget_type' not in session:
        return DEFAULT_BUDGET
    
    try:
        conn = sqlite3.connect(DB_NAME)
        cur = conn.cursor()
        cur.execute("SELECT budget FROM settings WHERE user_id = ? AND budget_type = ?", 
                   (session['user_id'], session['budget_type']))
        result = cur.fetchone()
        conn.close()
        return result[0] if result else DEFAULT_BUDGET
    except sqlite3.Error as e:
        print(f"Error fetching budget: {e}")
        return DEFAULT_BUDGET


def set_budget(amount):
    """Update the budget in settings"""
    if 'user_id' not in session or 'budget_type' not in session:
        return False
    
    try:
        conn = sqlite3.connect(DB_NAME)
        cur = conn.cursor()
        
        # Check if settings exist for user and budget type
        cur.execute("SELECT id FROM settings WHERE user_id = ? AND budget_type = ?", 
                   (session['user_id'], session['budget_type']))
        exists = cur.fetchone()
        
        if exists:
            cur.execute("UPDATE settings SET budget = ? WHERE user_id = ? AND budget_type = ?", 
                       (amount, session['user_id'], session['budget_type']))
        else:
            cur.execute("INSERT INTO settings (user_id, budget_type, budget) VALUES (?, ?, ?)", 
                       (session['user_id'], session['budget_type'], amount))
        
        conn.commit()
        conn.close()
        return True
    except sqlite3.Error as e:
        print(f"Error updating budget: {e}")
        return False


# Login required decorator
def login_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if 'user_id' not in session:
            return redirect(url_for('login_page'))
        return f(*args, **kwargs)
    return decorated_function


@app.route("/")
@login_required
def index():
    """Main route - show dashboard landing"""
    try:
        conn = sqlite3.connect(DB_NAME)
        cur = conn.cursor()
        cur.execute("""
            SELECT username, email, newsletter_subscribed, phone, email_alerts, sms_alerts 
            FROM users WHERE id = ?
        """, (session['user_id'],))
        user = cur.fetchone()
        conn.close()
        
        if user:
            return render_template("dashboard.html", 
                                 username=user[0], 
                                 email=user[1],
                                 newsletter_subscribed=user[2],
                                 phone=user[3] or "",
                                 email_alerts=user[4],
                                 sms_alerts=user[5])
        else:
            return render_template("dashboard.html", 
                                 username=session.get('username', 'User'), 
                                 email="",
                                 newsletter_subscribed=0,
                                 phone="",
                                 email_alerts=1,
                                 sms_alerts=0)
    except Exception as e:
        print(f"Error loading dashboard: {e}")
        return render_template("dashboard.html", 
                             username=session.get('username', 'User'), 
                             email="",
                             newsletter_subscribed=0,
                             phone="",
                             email_alerts=1,
                             sms_alerts=0)


@app.route("/tracker/<budget_type>")
@login_required
def tracker(budget_type):
    """Expense tracker for specific budget type"""
    if budget_type not in ['daily', 'weekly', 'monthly', 'annually']:
        return redirect(url_for('index'))
    
    session['budget_type'] = budget_type
    
    # Get budget type display name
    budget_names = {
        'daily': 'Daily',
        'weekly': 'Weekly',
        'monthly': 'Monthly',
        'annually': 'Annual'
    }
    
    return render_template("form.html", 
                         budget_type=budget_type,
                         budget_name=budget_names[budget_type])


@app.route("/update-profile", methods=["POST"])
@login_required
def update_profile():
    """Update user profile"""
    try:
        email = request.form.get("email", "").strip()
        phone = request.form.get("phone", "").strip()
        current_password = request.form.get("current_password", "")
        new_password = request.form.get("new_password", "")
        confirm_new_password = request.form.get("confirm_new_password", "")
        newsletter_subscription = request.form.get("newsletter_subscription") == "on"
        email_alerts = request.form.get("email_alerts") == "on"
        sms_alerts = request.form.get("sms_alerts") == "on"
        
        conn = sqlite3.connect(DB_NAME)
        cur = conn.cursor()
        
        # Get current user info
        cur.execute("""
            SELECT username, email, newsletter_subscribed 
            FROM users WHERE id = ?
        """, (session['user_id'],))
        user = cur.fetchone()
        current_newsletter_status = user[2] if user else 0
        
        # Update email and phone if changed
        if email:
            cur.execute("UPDATE users SET email = ? WHERE id = ?", (email, session['user_id']))
        
        if phone:
            cur.execute("UPDATE users SET phone = ? WHERE id = ?", (phone, session['user_id']))
        
        # Update newsletter subscription
        cur.execute("UPDATE users SET newsletter_subscribed = ? WHERE id = ?", 
                   (1 if newsletter_subscription else 0, session['user_id']))
        
        # Update alert preferences
        cur.execute("""
            UPDATE users 
            SET email_alerts = ?, sms_alerts = ? 
            WHERE id = ?
        """, (1 if email_alerts else 0, 1 if sms_alerts else 0, session['user_id']))
        
        # If newsletter status changed from unchecked to checked, add to spreadsheet
        if newsletter_subscription and not current_newsletter_status:
            username = user[0] if user else session.get('username', 'User')
            user_email = email if email else (user[1] if user else '')
            add_newsletter_subscriber(session['user_id'], username, user_email)
        
        # Update password if provided
        if current_password and new_password:
            # Verify current password
            cur.execute("SELECT password FROM users WHERE id = ?", (session['user_id'],))
            user = cur.fetchone()
            
            if user and check_password_hash(user[0], current_password):
                if new_password == confirm_new_password:
                    if len(new_password) >= 6:
                        hashed_password = generate_password_hash(new_password)
                        cur.execute("UPDATE users SET password = ? WHERE id = ?", (hashed_password, session['user_id']))
                        flash("Password updated successfully!", "success")
                    else:
                        flash("New password must be at least 6 characters", "error")
                        conn.close()
                        return redirect(url_for('index'))
                else:
                    flash("New passwords do not match", "error")
                    conn.close()
                    return redirect(url_for('index'))
            else:
                flash("Current password is incorrect", "error")
                conn.close()
                return redirect(url_for('index'))
        
        conn.commit()
        conn.close()
        
        flash("Profile updated successfully!", "success")
        return redirect(url_for('index'))
        
    except Exception as e:
        print(f"Error updating profile: {e}")
        flash("An error occurred while updating profile", "error")
        return redirect(url_for('index'))


@app.route("/signup", methods=["GET", "POST"])
def signup():
    """Sign up route"""
    if request.method == "POST":
        username = request.form.get("username", "").strip()
        email = request.form.get("email", "").strip()
        phone = request.form.get("phone", "").strip()
        password = request.form.get("password", "")
        confirm_password = request.form.get("confirm_password", "")
        newsletter = request.form.get("newsletter") == "on"
        sms_alerts = request.form.get("sms_alerts") == "on"
        
        print(f"Signup attempt - Username: {username}, Email: {email}, Phone: {phone}, Newsletter: {newsletter}, SMS: {sms_alerts}")
        
        # Validation
        if not username or not email or not password:
            flash("Username, email and password are required", "error")
            return redirect(url_for("signup"))
        
        if password != confirm_password:
            flash("Passwords do not match", "error")
            return redirect(url_for("signup"))
        
        if len(password) < 6:
            flash("Password must be at least 6 characters", "error")
            return redirect(url_for("signup"))
        
        # Validate phone if SMS alerts are enabled
        if sms_alerts and not phone:
            flash("Phone number is required for SMS alerts", "error")
            return redirect(url_for("signup"))
        
        try:
            conn = sqlite3.connect(DB_NAME)
            cur = conn.cursor()
            
            # Check if username or email already exists
            cur.execute("SELECT id FROM users WHERE username = ? OR email = ?", (username, email))
            if cur.fetchone():
                flash("Username or email already exists", "error")
                conn.close()
                return redirect(url_for("signup"))
            
            # Hash password and create user
            hashed_password = generate_password_hash(password)
            cur.execute("""
                INSERT INTO users (username, email, phone, password, newsletter_subscribed, sms_alerts, email_alerts) 
                VALUES (?, ?, ?, ?, ?, ?, ?)
            """, (username, email, phone if phone else None, hashed_password, 
                  1 if newsletter else 0, 1 if sms_alerts else 0, 1))
            user_id = cur.lastrowid
            
            # Create default settings for all budget types
            for budget_type in ['daily', 'weekly', 'monthly', 'annually']:
                cur.execute("INSERT INTO settings (user_id, budget_type, budget) VALUES (?, ?, ?)", 
                           (user_id, budget_type, DEFAULT_BUDGET))
            
            conn.commit()
            conn.close()
            
            # Add to newsletter spreadsheet if subscribed
            if newsletter:
                add_newsletter_subscriber(user_id, username, email)
            
            print(f"✓ User created successfully: {username} (ID: {user_id})")
            flash("Account created successfully! Please log in.", "success")
            return redirect(url_for("login_page"))
            
        except sqlite3.Error as e:
            print(f"✗ Database error during signup: {e}")
            flash(f"Database error: {str(e)}", "error")
            return redirect(url_for("signup"))
        except Exception as e:
            print(f"✗ Unexpected error during signup: {e}")
            flash(f"Error: {str(e)}", "error")
            return redirect(url_for("signup"))
    
    return render_template("signup.html")


@app.route("/login", methods=["GET", "POST"])
def login_page():
    """Login route"""
    if request.method == "POST":
        username = request.form.get("username", "").strip()
        password = request.form.get("password", "")
        
        if not username or not password:
            flash("Username and password are required", "error")
            return redirect(url_for("login_page"))
        
        try:
            conn = sqlite3.connect(DB_NAME)
            cur = conn.cursor()
            cur.execute("SELECT id, password, username FROM users WHERE username = ? OR email = ?", (username, username))
            user = cur.fetchone()
            conn.close()
            
            if user and check_password_hash(user[1], password):
                session['user_id'] = user[0]
                session['username'] = user[2]
                flash(f"Welcome back, {user[2]}!", "success")
                return redirect(url_for("index"))
            else:
                flash("Invalid username or password", "error")
                return redirect(url_for("login_page"))
                
        except sqlite3.Error as e:
            print(f"Database error: {e}")
            flash("An error occurred. Please try again.", "error")
            return redirect(url_for("login_page"))
    
    return render_template("login.html")


@app.route("/logout")
def logout():
    """Logout route"""
    username = session.get('username', 'User')
    session.clear()
    flash(f"Goodbye, {username}!", "success")
    return redirect(url_for("login_page"))


@app.route("/add_expense", methods=["POST"])
@login_required
def add_expense():
    """Add expense route"""
    if 'budget_type' not in session:
        return redirect(url_for('index'))
    
    try:
        category = request.form.get("category", "").strip()
        amount_str = request.form.get("amount", "")
        date = request.form.get("date", "")
        description = request.form.get("description", "").strip()

        # Validation
        if not category or not amount_str or not date:
            print("Error: Missing form data")
            return redirect(url_for('tracker', budget_type=session['budget_type']))

        try:
            amount = float(amount_str)
            if amount <= 0:
                print("Error: Amount must be greater than 0")
                return redirect(url_for('tracker', budget_type=session['budget_type']))
        except ValueError:
            print("Error: Invalid amount")
            return redirect(url_for('tracker', budget_type=session['budget_type']))

        # Validate date format
        try:
            datetime.strptime(date, "%Y-%m-%d")
        except ValueError:
            print("Error: Invalid date format")
            return redirect(url_for('tracker', budget_type=session['budget_type']))

        # Insert into database
        conn = sqlite3.connect(DB_NAME)
        cur = conn.cursor()
        cur.execute(
            "INSERT INTO expenses (user_id, budget_type, category, amount, date, description) VALUES (?, ?, ?, ?, ?, ?)",
            (session['user_id'], session['budget_type'], category, amount, date, description if description else None)
        )
        conn.commit()
        conn.close()
        
        desc_text = f" - {description}" if description else ""
        print(f"✓ Added to {session['budget_type']}: {category} - GH₵{amount} on {date}{desc_text}")
        
        # Check and send alerts after adding expense
        check_and_send_alerts(session['user_id'], session['budget_type'])
        
        return redirect(url_for('tracker', budget_type=session['budget_type']))
        
    except sqlite3.Error as e:
        print(f"Database error: {e}")
        return redirect(url_for('tracker', budget_type=session['budget_type']))
    except Exception as e:
        print(f"Error: {e}")
        return redirect(url_for('tracker', budget_type=session['budget_type']))


@app.route("/api/expenses", methods=["GET"])
@login_required
def api_expenses():
    """API endpoint to get all expenses and calculations"""
    expenses = get_expenses()
    budget = get_budget()
    
    # Calculate totals
    total_spent = sum(e[1] for e in expenses)
    remaining = budget - total_spent
    over_budget = total_spent > budget
    half_budget = budget / 2
    at_half_limit = total_spent >= half_budget and not over_budget
    
    # Calculate category totals
    category_totals = {}
    for cat, amt, _, _, _ in expenses:
        category_totals[cat] = category_totals.get(cat, 0) + amt
    
    # Format numbers
    total_spent = round(total_spent, 2)
    remaining = round(abs(remaining), 2)
    category_totals = {k: round(v, 2) for k, v in category_totals.items()}
    
    return jsonify({
        "expenses": expenses,
        "total_spent": total_spent,
        "remaining": remaining,
        "over_budget": over_budget,
        "at_half_limit": at_half_limit,
        "category_totals": category_totals,
        "budget": budget
    })


@app.route("/api/budget", methods=["POST"])
@login_required
def update_budget():
    """API endpoint to update the budget"""
    try:
        data = request.get_json()
        new_budget = float(data.get("budget", 0))
        
        if new_budget <= 0:
            return jsonify({"success": False, "message": "Budget must be greater than 0"}), 400
        
        if set_budget(new_budget):
            print(f"✓ Budget updated to GH₵{new_budget}")
            return jsonify({"success": True, "budget": new_budget})
        else:
            return jsonify({"success": False, "message": "Failed to update budget"}), 500
            
    except Exception as e:
        print(f"Error updating budget: {e}")
        return jsonify({"success": False, "message": str(e)}), 500


@app.route("/delete/<int:expense_id>", methods=["POST"])
@login_required
def delete_expense(expense_id):
    """Delete a specific expense by ID"""
    if 'budget_type' not in session:
        return redirect(url_for('index'))
    
    try:
        conn = sqlite3.connect(DB_NAME)
        cur = conn.cursor()
        
        # Check if expense exists
        cur.execute("SELECT category, amount FROM expenses WHERE id = ? AND user_id = ? AND budget_type = ?", 
                   (expense_id, session['user_id'], session['budget_type']))
        expense = cur.fetchone()
        
        if expense:
            # Delete the expense
            cur.execute("DELETE FROM expenses WHERE id = ?", (expense_id,))
            conn.commit()
            print(f"✓ Deleted from {session['budget_type']}: {expense[0]} - GH₵{expense[1]}")
        else:
            print("Error: Expense not found")
        
        conn.close()
        
    except sqlite3.Error as e:
        print(f"Error deleting expense: {e}")
    
    return redirect(url_for('tracker', budget_type=session['budget_type']))


@app.route("/edit/<int:expense_id>", methods=["POST"])
@login_required
def edit_expense(expense_id):
    """Edit a specific expense by ID"""
    if 'budget_type' not in session:
        return redirect(url_for('index'))
    
    try:
        category = request.form.get("category", "").strip()
        amount_str = request.form.get("amount", "")
        date = request.form.get("date", "")
        description = request.form.get("description", "").strip()

        # Validation
        if not category or not amount_str or not date:
            print("Error: Missing form data")
            return redirect(url_for('tracker', budget_type=session['budget_type']))

        try:
            amount = float(amount_str)
            if amount <= 0:
                print("Error: Amount must be greater than 0")
                return redirect(url_for('tracker', budget_type=session['budget_type']))
        except ValueError:
            print("Error: Invalid amount")
            return redirect(url_for('tracker', budget_type=session['budget_type']))

        # Validate date format
        try:
            datetime.strptime(date, "%Y-%m-%d")
        except ValueError:
            print("Error: Invalid date format")
            return redirect(url_for('tracker', budget_type=session['budget_type']))

        # Update in database
        conn = sqlite3.connect(DB_NAME)
        cur = conn.cursor()
        cur.execute(
            "UPDATE expenses SET category = ?, amount = ?, date = ?, description = ? WHERE id = ? AND user_id = ? AND budget_type = ?",
            (category, amount, date, description if description else None, expense_id, session['user_id'], session['budget_type'])
        )
        conn.commit()
        conn.close()
        
        desc_text = f" - {description}" if description else ""
        print(f"✓ Updated expense ID {expense_id} in {session['budget_type']}: {category} - GH₵{amount} on {date}{desc_text}")
        
    except sqlite3.Error as e:
        print(f"Error updating expense: {e}")
    except Exception as e:
        print(f"Error: {e}")
    
    return redirect(url_for('tracker', budget_type=session['budget_type']))


@app.route("/export/csv")
@login_required
def export_csv():
    """Export all expenses to CSV file"""
    try:
        expenses = get_expenses()
        budget = get_budget()
        
        # Create CSV in memory
        output = io.StringIO()
        writer = csv.writer(output)
        
        # Write header
        writer.writerow(['Category', 'Amount (GH₵)', 'Date', 'Description', 'ID'])
        
        # Write expense data
        for exp in expenses:
            writer.writerow([exp[0], exp[1], exp[2], exp[4] if exp[4] else '', exp[3]])
        
        # Add summary at the end
        writer.writerow([])
        writer.writerow(['Summary'])
        writer.writerow(['Total Spent', sum(e[1] for e in expenses)])
        writer.writerow(['Budget', budget])
        writer.writerow(['Remaining', budget - sum(e[1] for e in expenses)])
        
        # Prepare file for download
        output.seek(0)
        
        # Create filename with current date
        filename = f"rahmonapay_expenses_{datetime.now().strftime('%Y%m%d')}.csv"
        
        # Convert to bytes
        mem = io.BytesIO()
        mem.write(output.getvalue().encode('utf-8'))
        mem.seek(0)
        
        print(f"✓ Exported {len(expenses)} expenses to CSV")
        
        return send_file(
            mem,
            mimetype='text/csv',
            as_attachment=True,
            download_name=filename
        )
        
    except Exception as e:
        print(f"Error exporting CSV: {e}")
        return redirect(url_for("index"))


@app.route("/export/excel")
@login_required
def export_excel():
    """Export all expenses to Excel-compatible CSV"""
    try:
        expenses = get_expenses()
        budget = get_budget()
        
        # Create CSV in memory (Excel can open CSV files)
        output = io.StringIO()
        writer = csv.writer(output)
        
        # Write title
        writer.writerow(['RahmonaPay Expense Report'])
        writer.writerow([f'Generated: {datetime.now().strftime("%Y-%m-%d %H:%M")}'])
        writer.writerow([])
        
        # Write header
        writer.writerow(['Category', 'Amount (GH₵)', 'Date', 'Description', 'Expense ID'])
        
        # Write expense data
        for exp in expenses:
            writer.writerow([exp[0], f'{exp[1]:.2f}', exp[2], exp[4] if exp[4] else '', exp[3]])
        
        # Add summary section
        writer.writerow([])
        writer.writerow(['SUMMARY'])
        writer.writerow(['Monthly Budget', f'{budget:.2f}'])
        writer.writerow(['Total Spent', f'{sum(e[1] for e in expenses):.2f}'])
        writer.writerow(['Remaining', f'{budget - sum(e[1] for e in expenses):.2f}'])
        
        # Category breakdown
        writer.writerow([])
        writer.writerow(['CATEGORY BREAKDOWN'])
        category_totals = {}
        for cat, amt, _, _, _ in expenses:
            category_totals[cat] = category_totals.get(cat, 0) + amt
        
        for cat, amt in category_totals.items():
            writer.writerow([cat, f'{amt:.2f}'])
        
        # Prepare file for download
        output.seek(0)
        
        # Create filename with current date
        filename = f"rahmonapay_report_{datetime.now().strftime('%Y%m%d')}.csv"
        
        # Convert to bytes
        mem = io.BytesIO()
        mem.write(output.getvalue().encode('utf-8'))
        mem.seek(0)
        
        print(f"✓ Exported {len(expenses)} expenses to Excel format")
        
        return send_file(
            mem,
            mimetype='text/csv',
            as_attachment=True,
            download_name=filename
        )
        
    except Exception as e:
        print(f"Error exporting Excel: {e}")
        return redirect(url_for("index"))


@app.route("/download-newsletter-list")
@login_required
def download_newsletter_list():
    """Download the newsletter subscribers Excel file (admin only)"""
    try:
        if not os.path.exists(NEWSLETTER_FILE):
            flash("Newsletter file not found", "error")
            return redirect(url_for("index"))
        
        return send_file(
            NEWSLETTER_FILE,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name=f"newsletter_subscribers_{datetime.now().strftime('%Y%m%d')}.xlsx"
        )
    except Exception as e:
        print(f"Error downloading newsletter list: {e}")
        flash("Error downloading newsletter list", "error")
        return redirect(url_for("index"))


@app.route("/api/dashboard-stats")
@login_required
def dashboard_stats():
    """API endpoint to get dashboard statistics"""
    try:
        conn = sqlite3.connect(DB_NAME)
        cur = conn.cursor()
        
        # Get all user expenses across all budget types
        cur.execute("""
            SELECT category, amount, date, budget_type 
            FROM expenses 
            WHERE user_id = ? 
            ORDER BY date DESC
        """, (session['user_id'],))
        all_expenses = cur.fetchall()
        
        # Get all budgets
        cur.execute("""
            SELECT budget_type, budget 
            FROM settings 
            WHERE user_id = ?
        """, (session['user_id'],))
        budgets = dict(cur.fetchall())
        
        conn.close()
        
        # Calculate total expenses
        total_expenses = sum(exp[1] for exp in all_expenses)
        
        # Calculate week expenses (last 7 days)
        today = datetime.now()
        week_ago = today - timedelta(days=7)
        week_expenses = sum(
            exp[1] for exp in all_expenses 
            if datetime.strptime(exp[2], "%Y-%m-%d") >= week_ago
        )
        
        # Calculate month expenses (current month)
        current_month = today.strftime("%Y-%m")
        month_expenses = sum(
            exp[1] for exp in all_expenses 
            if exp[2].startswith(current_month)
        )
        
        # Calculate budget status
        total_budget = sum(budgets.values()) if budgets else 0
        budget_status = 'good'
        
        if total_budget > 0:
            if total_expenses > total_budget:
                budget_status = 'over'
            elif total_expenses >= total_budget * 0.8:
                budget_status = 'warning'
        
        # Category totals
        category_totals = {}
        for cat, amt, _, _ in all_expenses:
            category_totals[cat] = category_totals.get(cat, 0) + amt
        
        # Recent expenses for activity feed
        recent_expenses = [
            {
                'category': exp[0],
                'amount': exp[1],
                'date': exp[2],
                'budget_type': exp[3]
            }
            for exp in all_expenses[:10]
        ]
        
        return jsonify({
            'total_expenses': round(total_expenses, 2),
            'week_expenses': round(week_expenses, 2),
            'month_expenses': round(month_expenses, 2),
            'budget_status': budget_status,
            'category_totals': {k: round(v, 2) for k, v in category_totals.items()},
            'recent_expenses': recent_expenses
        })
        
    except Exception as e:
        print(f"Error getting dashboard stats: {e}")
        return jsonify({
            'total_expenses': 0,
            'week_expenses': 0,
            'month_expenses': 0,
            'budget_status': 'good',
            'category_totals': {},
            'recent_expenses': []
        })


@app.route("/init-database")
def initialize_database():
    """Manual database initialization endpoint"""
    try:
        init_db()
        init_newsletter_spreadsheet()
        return """
        <html>
        <head>
            <title>Database Initialized</title>
            <style>
                body {
                    font-family: system-ui, sans-serif;
                    background: #0f172a;
                    color: #e5e7eb;
                    display: flex;
                    align-items: center;
                    justify-content: center;
                    height: 100vh;
                    margin: 0;
                }
                .message {
                    text-align: center;
                    padding: 40px;
                    background: #1e293b;
                    border-radius: 20px;
                    border: 2px solid #38bdf8;
                }
                h1 { color: #38bdf8; margin-bottom: 20px; }
                a {
                    display: inline-block;
                    margin-top: 20px;
                    padding: 12px 24px;
                    background: linear-gradient(135deg, #38bdf8, #2563eb);
                    color: white;
                    text-decoration: none;
                    border-radius: 10px;
                    font-weight: 600;
                }
                a:hover { transform: translateY(-2px); }
            </style>
        </head>
        <body>
            <div class="message">
                <h1>✓ Database Initialized Successfully!</h1>
                <p>Your RahmonaPay database has been set up and is ready to use.</p>
                <a href="/login">Go to Login Page</a>
            </div>
        </body>
        </html>
        """
    except Exception as e:
        return f"""
        <html>
        <head>
            <title>Database Error</title>
            <style>
                body {{
                    font-family: system-ui, sans-serif;
                    background: #0f172a;
                    color: #e5e7eb;
                    display: flex;
                    align-items: center;
                    justify-content: center;
                    height: 100vh;
                    margin: 0;
                }}
                .message {{
                    text-align: center;
                    padding: 40px;
                    background: #1e293b;
                    border-radius: 20px;
                    border: 2px solid #ef4444;
                }}
                h1 {{ color: #ef4444; margin-bottom: 20px; }}
            </style>
        </head>
        <body>
            <div class="message">
                <h1>✗ Error Initializing Database</h1>
                <p>Error: {str(e)}</p>
            </div>
        </body>
        </html>
        """


if __name__ == "__main__":
    init_db()
    init_newsletter_spreadsheet()
    print("=" * 60)
    print("🚀 RahmonaPay Expense Tracker Starting...")
    print("=" * 60)
    print("📊 IMPORTANT: Open this URL in your browser:")
    print("   👉 http://127.0.0.1:5000")
    print("=" * 60)
    print("⚠️  DO NOT open form.html directly from your folders!")
    print("   You MUST use the URL above for the app to work.")
    print("=" * 60)
    print(f"📧 Newsletter subscribers will be saved to: {NEWSLETTER_FILE}")
    print("=" * 60)
    print("Press Ctrl+C to stop the server")
    print("=" * 60)
    app.run(debug=True, host='127.0.0.1', port=5000)