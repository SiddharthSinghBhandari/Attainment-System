from pymongo import MongoClient
import streamlit as st

# Load MongoDB URI securely from secrets.toml
MONGO_URI = st.secrets["mongo_uri"]

# Create MongoDB client
client = MongoClient(MONGO_URI)

# Select database
db = client["attainment_system"]

# Collections
students = db["students"]
users = db["users"]

# Collection for uploaded Excel data
uploads = db["teacher_uploads"]