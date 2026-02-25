from pymongo import MongoClient
import streamlit as st

MONGO_URI = st.secrets["mongo_uri"]

client = MongoClient(MONGO_URI)

db = client["attainment_system"]

students = db["students"]