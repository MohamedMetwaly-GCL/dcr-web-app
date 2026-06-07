import sys
from app import app
import db

with app.test_request_context('/api/daily_digest/all'):
    try:
        from auth import get_allowed_project_ids
        projects_to_check = [p["id"] for p in db.get_projects()] # mock
        combined = {"received": [], "issued": [], "replied": []}
        for p_id in projects_to_check:
            dist = db.get_distribution(p_id)
            d = db.get_daily_digest(p_id, ["LTR", "DS"])
            combined["received"].extend(d["received"])
        print(combined)
    except Exception as e:
        import traceback
        traceback.print_exc()
