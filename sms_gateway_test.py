import requests
import json
import logging

# Set up logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    filename='sms_api_test.log'
)
logger = logging.getLogger(__name__)

def send_sms_via_api(phone_number, message, api_key):
    """
    Send SMS via SMS.to API
    
    Args:
        phone_number (str): Recipient's phone number with country code
        message (str): Message content
        api_key (str): Your SMS.to API key
    
    Returns:
        bool: True if successful, False otherwise
    """
    url = "https://api.sms.to/sms/send"
    
    headers = {
        "Authorization": f"Bearer {api_key}",
        "Content-Type": "application/json"
    }
    
    payload = {
        "message": message,
        "to": phone_number,
        "sender_id": "Schedule"
    }
    
    try:
        logger.info(f"Attempting to send SMS to: {phone_number}")
        response = requests.post(url, headers=headers, data=json.dumps(payload))
        response_data = response.json()
        
        logger.info(f"API Response: {response_data}")
        
        if response.status_code >= 200 and response.status_code < 300:
            logger.info(f"Successfully sent SMS. Response: {response.text}")
            print(f"Message sent successfully to {phone_number}")
            return True
        else:
            error_message = f"Failed to send SMS. Status code: {response.status_code}, Response: {response.text}"
            logger.error(error_message)
            print(error_message)
            return False
        
    except requests.exceptions.RequestException as e:
        error_message = f"Failed to send SMS: {str(e)}"
        logger.error(error_message)
        print(error_message)
        return False

def main():
    """Interactive test function for SMS.to API"""
    print("SMS.to API Test")
    print("-" * 30)
    
    # Your provided information
    phone_number = "+85266631823"  # Your test phone number with country code
    api_key = "eyJ0eXAiOiJKV1QiLCJhbGciOiJIUzI1NiJ9.eyJpc3MiOiJodHRwczovL2F1dGg6ODA4MC9hcGkvdjEvdXNlcnMvYXBpL2tleXMvZ2VuZXJhdGUiLCJpYXQiOjE3NDQzNTI2MjAsIm5iZiI6MTc0NDM1MjYyMCwianRpIjoiZnkxQUp2bTBFMkM3d1lCZyIsInN1YiI6NDc4MDcyLCJwcnYiOiIyM2JkNWM4OTQ5ZjYwMGFkYjM5ZTcwMWM0MDA4NzJkYjdhNTk3NmY3In0.sG53Ml7UUu6OfgtNZoQySfeJBrtk2Q1k8yZYVrLrwAw"
    
    # Get message
    message = input("Enter your test message: ")
    
    # Send the message
    print("\nSending test SMS...")
    result = send_sms_via_api(phone_number, message, api_key)
    
    if result:
        print("\nTest completed successfully! Check your phone for the message.")
    else:
        print("\nTest failed. Check the log file for details.")
        print("Note: Make sure you have sufficient credits in your SMS.to account.")

if __name__ == "__main__":
    main()
