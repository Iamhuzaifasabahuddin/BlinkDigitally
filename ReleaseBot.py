import time
import streamlit as st
from slack_sdk import WebClient
from slack_sdk.errors import SlackApiError

SLACK_BOT_TOKEN = st.secrets["Slack"]["Slack"]
client = WebClient(token=SLACK_BOT_TOKEN)

US_CHANNEL_ID = st.secrets["Channels"]["usa"]
UK_CHANNEL_ID = st.secrets["Channels"]["uk"]
APP_PASSWORD = st.secrets["BOT"]["password"]

st.set_page_config(page_title="Slack Version Bot", page_icon="🤖", layout="centered")

def get_user_id_by_email(email: str):
    """Retrieve Slack user ID from email address."""
    try:
        response = client.users_lookupByEmail(email=email)
        return response['user']['id']
    except SlackApiError as e:
        st.error(f"Failed to get user ID: {e.response['error']}")
        return None


def send_dm(user_id: str, message: str):
    """Send a direct message (DM) to a user by user ID."""
    try:
        client.chat_postMessage(channel=user_id, text=message)
        return True
    except SlackApiError as e:
        st.error(f"Error sending DM: {e.response['error']}")
        return False


def send_to_channel(channel_id: str, message: str):
    """Send a message to a Slack channel by channel ID."""
    try:
        client.chat_postMessage(channel=channel_id, text=message)
        return True
    except SlackApiError as e:
        st.error(f"Error sending to channel: {e.response['error']}")
        return False

if "authenticated" not in st.session_state:
    st.session_state.authenticated = False

if not st.session_state.authenticated:
    st.title("🔑 Version Release Bot Login")
    password = st.text_input("Enter Password", type="password")
    if st.button("Login"):
        if password == APP_PASSWORD:
            st.session_state.authenticated = True
            st.success("Login successful! 🎉")
            time.sleep(2)
            st.rerun()
        else:
            st.error("Incorrect password.")
    st.stop()

st.title("🤖 Version Update Bot")
st.markdown("Send version updates directly to **US**, **UK**, **Both**, or as a **Hexz DM**.")


with st.form("version_form", clear_on_submit=False):
    st.subheader("📝 Version Update Details")

    region = st.selectbox("Select Region:", ["USA", "UK", "Both", "Hexz"])
    description = st.text_area("Enter Description:")
    version_name = st.text_input("Enter Version Name:")

    submitted = st.form_submit_button("🚀 Send Update")

    if submitted:
        if not description or not version_name:
            st.warning("⚠️ Please fill in both the description and version name.")
        else:
            message = (
                f"*📦 Version Update!*\n\n"
                f"🚀 *Version:* {version_name}\n"
                f"📄 *Description:*\n{description}"
            )

            success_list, error_list = [], []

            if region == "USA":
                if send_to_channel(US_CHANNEL_ID, message):
                    success_list.append("USA Channel")
            elif region == "UK":
                if send_to_channel(UK_CHANNEL_ID, message):
                    success_list.append("UK Channel")
            elif region == "Both":
                if send_to_channel(US_CHANNEL_ID, message):
                    success_list.append("USA Channel")
                if send_to_channel(UK_CHANNEL_ID, message):
                    success_list.append("UK Channel")
            elif region == "Hexz":

                user_id = get_user_id_by_email("huzaifa.sabah@topsoftdigitals.pk")
                if user_id and send_dm(user_id, message):
                    success_list.append("Hexz (Direct Message)")

            if success_list:
                st.success(f"✅ Message successfully sent to: {', '.join(success_list)}")
            else:
                st.error("⚠️ Failed to send message. Please check logs or email address.")
