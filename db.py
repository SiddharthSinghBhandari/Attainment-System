from pymongo import MongoClient
import os

MONGO_URI = os.environ.get("MONGO_URI")

client = MongoClient(MONGO_URI)

db = client["attainment_system"]

students = db["students"]
users = db["users"]
uploads = db["teacher_uploads"]