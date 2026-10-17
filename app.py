from flask import Flask, render_template_string, request, jsonify

app = Flask(__name__)

# Rider App Configuration
APP_NAME = "Rider"
OWNER_NAME = "Sonu Kumar"

# Delhi NCR Popular Locations
DELHI_NCR_LOCATIONS = [
    "Connaught Place, New Delhi",
    "Cyber Hub, Gurugram",
    "Sector 18, Noida",
    "Hauz Khas Village, New Delhi",
    "Indirapuram, Ghaziabad",
    "Chandni Chowk, Old Delhi",
    "DLF Phase 3, Gurugram",
    "Noida Electronic City, Sector 62",
    "Saket District Centre, New Delhi",
    "Rajiv Chowk Metro Station, Delhi"
]

# Available Rider Captains in Delhi NCR
CAPTAINS = [
    {"name": "Rahul Sharma", "bike": "Honda Shine (DL 3S AB 1234)", "rating": 4.9, "distance": "0.8 km away"},
    {"name": "Amit Kumar", "bike": "Hero Splendor (UP 16 X 5678)", "rating": 4.8, "distance": "1.2 km away"},
    {"name": "Vikas Verma", "bike": "TVS Apache (HR 26 C 9012)", "rating": 4.9, "distance": "0.5 km away"}
]

# HTML & CSS UI (English Only)
HTML_TEMPLATE = """
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Rider - Quick Bike Rides in Delhi NCR</title>
    <style>
        * { box-sizing: border-box; margin: 0; padding: 0; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; }
        body { background-color: #f4f6f8; color: #333; }
        
        header { background-color: #ffcc00; padding: 15px 20px; display: flex; justify-content: space-between; align-items: center; box-shadow: 0 2px 8px rgba(0,0,0,0.1); }
        .logo { font-size: 24px; font-weight: bold; color: #111; }
        .owner-badge { background: #111; color: #ffcc00; padding: 5px 12px; border-radius: 20px; font-size: 12px; font-weight: 600; }
        
        .container { max-width: 600px; margin: 30px auto; padding: 20px; }
        .card { background: white; border-radius: 12px; padding: 24px; box-shadow: 0 4px 12px rgba(0,0,0,0.08); margin-bottom: 20px; }
        
        h2 { font-size: 20px; margin-bottom: 15px; color: #111; }
        label { font-size: 14px; font-weight: 600; color: #555; display: block; margin-top: 12px; margin-bottom: 5px; }
        
        select, input { width: 100%; padding: 12px; border: 1px solid #ccc; border-radius: 8px; font-size: 15px; outline: none; }
        select:focus, input:focus { border-color: #ffcc00; }
        
        .btn { width: 100%; background: #ffcc00; color: #111; border: none; padding: 14px; font-size: 16px; font-weight: bold; border-radius: 8px; cursor: pointer; margin-top: 20px; transition: 0.2s; }
        .btn:hover { background: #e6b800; }
        
        .ride-option { display: flex; justify-content: space-between; align-items: center; border: 2px solid #eee; border-radius: 8px; padding: 12px; margin-top: 10px; cursor: pointer; }
        .ride-option.active { border-color: #ffcc00; background-color: #fffdf0; }
        
        .captain-card { border-left: 4px solid #ffcc00; background: #fafafa; padding: 12px; margin-top: 10px; border-radius: 4px; }
        footer { text-align: center; margin-top: 40px; font-size: 13px; color: #777; }
    </style>
</head>
<body>

    <header>
        <div class="logo">⚡ RIDER</div>
        <div class="owner-badge">Owner: {{ owner }}</div>
    </header>

    <div class="container">
        <div class="card">
            <h2>Book a Bike Ride in Delhi NCR</h2>
            
            <form action="/book" method="POST">
                <label for="pickup">Pickup Location:</label>
                <select name="pickup" id="pickup" required>
                    {% for loc in locations %}
                        <option value="{{ loc }}">{{ loc }}</option>
                    {% endfor %}
                </select>

                <label for="drop">Drop Location:</label>
                <select name="drop" id="drop" required>
                    {% for loc in locations[::-1] %}
                        <option value="{{ loc }}">{{ loc }}</option>
                    {% endfor %}
                </select>

                <label>Select Vehicle Type:</label>
                <div class="ride-option active">
                    <div>
                        <strong>🛵 Rider Bike</strong>
                        <div style="font-size: 12px; color: #666;">Fastest through Delhi NCR traffic</div>
                    </div>
                    <div><strong>₹65</strong></div>
                </div>

                <button type="submit" class="btn">Find Nearest Rider Captain</button>
            </form>
        </div>

        {% if booking %}
        <div class="card" style="border-top: 4px solid #28a745;">
            <h2 style="color: #28a745;">✅ Ride Confirmed!</h2>
            <p><strong>From:</strong> {{ booking.pickup }}</p>
            <p><strong>To:</strong> {{ booking.drop }}</p>
            <p><strong>Estimated Fare:</strong> ₹65 (Cash / UPI)</p>
            
            <hr style="margin: 15px 0; border: 0; border-top: 1px solid #eee;">
            
            <h3>Assigned Rider Captain:</h3>
            <div class="captain-card">
                <strong>👤 {{ captain.name }}</strong> (⭐ {{ captain.rating }})<br>
                <span style="font-size: 13px; color: #555;">Bike: {{ captain.bike }}</span><br>
                <span style="font-size: 13px; color: #28a745; font-weight: bold;">Arriving in 3 mins ({{ captain.distance }})</span>
            </div>
        </div>
        {% endif %}

        <footer>
            &copy; 2026 Rider Inc. Owned & Managed by {{ owner }}. All rights reserved.
        </footer>
    </div>

</body>
</html>
"""

@app.route("/")
def home():
    return render_template_string(
        HTML_TEMPLATE, 
        owner=OWNER_NAME, 
        locations=DELHI_NCR_LOCATIONS, 
        booking=None
    )

@app.route("/book", methods=["POST"])
def book_ride():
    pickup = request.form.get("pickup")
    drop = request.form.get("drop")
    
    # Assign a random captain from Delhi NCR
    import random
    assigned_captain = random.choice(CAPTAINS)
    
    booking_details = {
        "pickup": pickup,
        "drop": drop
    }
    
    return render_template_string(
        HTML_TEMPLATE, 
        owner=OWNER_NAME, 
        locations=DELHI_NCR_LOCATIONS, 
        booking=booking_details,
        captain=assigned_captain
    )

if __name__ == "__main__":
    app.run(debug=True, port=5000)