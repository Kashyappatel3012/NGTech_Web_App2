"""
Daily Workplan Email Scheduler
Sends daily workplan Excel file via email at 12:05 PM IST
"""
import os
import logging
from datetime import datetime
from flask_mail import Message
from pytz import timezone
from apscheduler.schedulers.background import BackgroundScheduler
from apscheduler.triggers.cron import CronTrigger

# Configure logging
logger = logging.getLogger(__name__)

# IST timezone
ist = timezone('Asia/Kolkata')

def get_month_name(month):
    """Convert month number to month name"""
    month_names = {
        1: 'Jan', 2: 'Feb', 3: 'Mar', 4: 'Apr',
        5: 'May', 6: 'Jun', 7: 'Jul', 8: 'Aug',
        9: 'Sep', 10: 'Oct', 11: 'Nov', 12: 'Dec'
    }
    return month_names.get(month, 'Jan')

def generate_filename():
    """Generate filename based on current date in IST"""
    now = datetime.now(ist)
    day = str(now.day)
    month = get_month_name(now.month)
    year = str(now.year)
    return f"{day}_{month}_{year}.xlsx"

def send_daily_workplan_email():
    """Send daily workplan Excel file via email at 12:05 PM IST"""
    try:
        # Import here to avoid circular imports
        from app import mail, app
        
        with app.app_context():
            # Generate filename for today's date
            filename = generate_filename()
            file_path = os.path.join('static', 'Activity_Tracker', 'Everyday_Workplan', filename)
            
            # Recipient email
            recipient_email = 'ngt-auakua@ngtech.co.in'
            
            # Check if file exists
            if os.path.exists(file_path):
                # File exists - send email with attachment
                try:
                    msg = Message(
                        subject=f'Daily Workplan - {filename}',
                        recipients=[recipient_email],
                        body=f'Please find attached the daily workplan file: {filename}'
                    )
                    
                    # Attach the Excel file
                    with open(file_path, 'rb') as f:
                        msg.attach(
                            filename,
                            'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                            f.read()
                        )
                    
                    mail.send(msg)
                    logger.info(f"Successfully sent daily workplan email with file: {filename} to {recipient_email}")
                    print(f"✅ Successfully sent daily workplan email with file: {filename} to {recipient_email}")
                    
                except Exception as e:
                    logger.error(f"Error sending email with attachment: {str(e)}")
                    print(f"❌ Error sending email with attachment: {str(e)}")
                    
            else:
                # File not found - send email with "file not found" message
                try:
                    msg = Message(
                        subject=f'Daily Workplan - File Not Found',
                        recipients=[recipient_email],
                        body=f'File not found: {filename}\n\nExpected file path: {file_path}'
                    )
                    
                    mail.send(msg)
                    logger.warning(f"Sent email notification: File not found - {filename} to {recipient_email}")
                    print(f"⚠️ Sent email notification: File not found - {filename} to {recipient_email}")
                    
                except Exception as e:
                    logger.error(f"Error sending 'file not found' email: {str(e)}")
                    print(f"❌ Error sending 'file not found' email: {str(e)}")
                    
    except Exception as e:
        logger.error(f"Error in send_daily_workplan_email: {str(e)}")
        print(f"❌ Error in send_daily_workplan_email: {str(e)}")
        import traceback
        traceback.print_exc()

def send_daily_updated_work_email():
    """Send daily updated work Excel file via email at 11:05 PM IST"""
    try:
        # Import here to avoid circular imports
        from app import mail, app
        
        with app.app_context():
            # Generate filename for today's date
            filename = generate_filename()
            file_path = os.path.join('static', 'Activity_Tracker', 'Everyday_Updated_Work', filename)
            
            # Recipient email
            recipient_email = 'ngt-auakua@ngtech.co.in'
            
            # Check if file exists
            if os.path.exists(file_path):
                # File exists - send email with attachment
                try:
                    msg = Message(
                        subject=f'Daily Updated Work - {filename}',
                        recipients=[recipient_email],
                        body=f'Please find attached the daily updated work file: {filename}'
                    )
                    
                    # Attach the Excel file
                    with open(file_path, 'rb') as f:
                        msg.attach(
                            filename,
                            'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                            f.read()
                        )
                    
                    mail.send(msg)
                    logger.info(f"Successfully sent daily updated work email with file: {filename} to {recipient_email}")
                    print(f"✅ Successfully sent daily updated work email with file: {filename} to {recipient_email}")
                    
                except Exception as e:
                    logger.error(f"Error sending email with attachment: {str(e)}")
                    print(f"❌ Error sending email with attachment: {str(e)}")
                    
            else:
                # File not found - send email with "file not found" message
                try:
                    msg = Message(
                        subject=f'Daily Updated Work - File Not Found',
                        recipients=[recipient_email],
                        body=f'File not found: {filename}\n\nExpected file path: {file_path}'
                    )
                    
                    mail.send(msg)
                    logger.warning(f"Sent email notification: File not found - {filename} to {recipient_email}")
                    print(f"⚠️ Sent email notification: File not found - {filename} to {recipient_email}")
                    
                except Exception as e:
                    logger.error(f"Error sending 'file not found' email: {str(e)}")
                    print(f"❌ Error sending 'file not found' email: {str(e)}")
                    
    except Exception as e:
        logger.error(f"Error in send_daily_updated_work_email: {str(e)}")
        print(f"❌ Error in send_daily_updated_work_email: {str(e)}")
        import traceback
        traceback.print_exc()

def start_scheduler():
    """Start the background scheduler for daily emails"""
    scheduler = BackgroundScheduler(timezone=ist)
    
    # Schedule workplan email to be sent every day at 12:05 PM IST
    scheduler.add_job(
        func=send_daily_workplan_email,
        trigger=CronTrigger(hour=12, minute=5, timezone=ist),
        id='daily_workplan_email',
        name='Send daily workplan email at 12:05 PM IST',
        replace_existing=True
    )
    
    # Schedule updated work email to be sent every day at 11:05 PM IST
    scheduler.add_job(
        func=send_daily_updated_work_email,
        trigger=CronTrigger(hour=23, minute=5, timezone=ist),
        id='daily_updated_work_email',
        name='Send daily updated work email at 11:05 PM IST',
        replace_existing=True
    )
    
    scheduler.start()
    logger.info("Daily email scheduler started:")
    logger.info("   - Workplan email: every day at 12:05 PM IST")
    logger.info("   - Updated work email: every day at 11:05 PM IST")
    print("📧 Daily email scheduler started:")
    print("   - Workplan email: every day at 12:05 PM IST")
    print("   - Updated work email: every day at 11:05 PM IST")
    
    return scheduler

