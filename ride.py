import os
import random
import urllib.parse
import requests
from flask import Flask, render_template_string, request, jsonify, redirect, session
from openpyxl import Workbook, load_workbook

app = Flask(__name__)
app.secret_key = "rider_secret_key_sonu_5067"  # Session security key

APP_NAME = "Rider"
OWNER_NAME = "Sonu Kumar"
OWNER_WHATSAPP_NUMBER = "917294182077"
OWNER_UPI_ID = "ssingh7479555065@oksbi"
ADMIN_PASSWORD = "sonu5067"
CAPTAIN_PASSWORD = "captain123"  # 👈 Captain ka Secret Password
EXCEL_FILE = "rider_bookings.xlsx"
CAPTAINS_FILE = "registered_captains.xlsx"

USER_OTP_STORE = {}

VALID_COUPONS = {
    "RIDER50": 10,
    "SONU20": 20,
    "FREEBIE": 15
}

PRESET_LANDMARKS = [
    {"name": "GNIOT College, Knowledge Park II, Greater Noida", "lat": 28.4607, "lon": 77.4912},
    {"name": "KCC Institute of Technology & Management, Knowledge Park III, Greater Noida", "lat": 28.4632, "lon": 77.4948},
    {"name": "Noida Electronic City Metro Station, Sector 62, Noida", "lat": 28.6276, "lon": 77.3725},
    {"name": "Knowledge Park III, Greater Noida", "lat": 28.4650, "lon": 77.4930},
    {"name": "Pari Chowk, Greater Noida", "lat": 28.4670, "lon": 77.5138},
    {"name": "Chatro, Deori, Giridih, Jharkhand", "lat": 24.5683, "lon": 86.2731}
]

def init_excel():
    if not os.path.exists(EXCEL_FILE):
        wb = Workbook()
        ws = wb.active
        ws.title = "Bookings"
        ws.append(["Customer Name", "Phone", "Pickup", "Drop", "Service", "Fare", "Payment Mode", "Captain Assigned", "Status", "Ride OTP", "Lat", "Lng", "Rating", "Coupon Used"])
        wb.save(EXCEL_FILE)

    if not os.path.exists(CAPTAINS_FILE):
        wb = Workbook()
        ws = wb.active
        ws.title = "Captains"
        ws.append(["Captain Name", "Phone", "Bike Model", "Vehicle Number", "Licence / Aadhaar", "Status", "Total Earnings", "Rating"])
        ws.append(["Mukesh Kumar", "9876543210", "Hero Splendor", "JH 10 AB 1234", "DL-987654321", "Approved", "₹0", "4.9 ⭐"])
        ws.append(["Santosh Mahto", "9123456780", "Honda Shine", "JH 10 CD 5678", "DL-123456789", "Approved", "₹0", "4.8 ⭐"])
        wb.save(CAPTAINS_FILE)

def save_to_excel(name, phone, pickup, drop, service, fare, payment_mode, captain, otp, lat, lng, coupon):
    init_excel()
    wb = load_workbook(EXCEL_FILE)
    ws = wb["Bookings"]
    ws.append([name, phone, pickup, drop, service, fare, payment_mode, captain, "Pending", otp, lat, lng, "Not Rated", coupon])
    wb.save(EXCEL_FILE)
    return ws.max_row - 2

LOGIN_FORM_HTML = """
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Login Access - Rider</title>
    <style>
        * { box-sizing: border-box; margin: 0; padding: 0; font-family: system-ui, -apple-system, sans-serif; }
        body { background-color: #030712; color: #f9fafb; display: flex; justify-content: center; align-items: center; min-height: 100vh; padding: 15px; }
        .login-card { background: #111827; border: 1px solid rgba(255,255,255,0.1); border-radius: 16px; padding: 25px; width: 100%; max-width: 380px; text-align: center; }
        h2 { color: #fbbf24; margin-bottom: 8px; }
        p { font-size: 13px; color: #9ca3af; margin-bottom: 20px; }
        input { width: 100%; padding: 12px; background: #030712; border: 1px solid #374151; border-radius: 8px; color: white; margin-bottom: 15px; font-size: 14px; outline: none; }
        input:focus { border-color: #fbbf24; }
        button { width: 100%; padding: 12px; background: #fbbf24; color: black; font-weight: bold; border: none; border-radius: 8px; cursor: pointer; font-size: 14px; }
        .error { color: #f87171; font-size: 12px; margin-bottom: 12px; }
        a { color: #9ca3af; text-decoration: none; font-size: 12px; display: inline-block; margin-top: 15px; }
    </style>
</head>
<body>
    <div class="login-card">
        <h2>🔒 {{ role }} Authentication</h2>
        <p>Restricted Area - Only Authorized Personnel</p>
        
        {% if error %}
        <div class="error">❌ {{ error }}</div>
        {% endif %}

        <form method="POST">
            {% if role == 'Captain' %}
            <input type="tel" name="phone" placeholder="Enter Registered Mobile Number" required maxlength="10">
            {% else %}
            <input type="password" name="password" placeholder="Enter Admin Secret Password" required>
            {% endif %}
            <button type="submit">Unlock Dashboard 🔑</button>
        </form>

        <a href="/">⬅ Return to Customer App</a>
    </div>
</body>
</html>
"""

HTML_TEMPLATE = """
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Rider - Live Ride Hailing</title>
    <link rel="stylesheet" href="https://unpkg.com/leaflet@1.9.4/dist/leaflet.css" />
    <style>
        * { box-sizing: border-box; margin: 0; padding: 0; font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif; }
        body { background-color: #030712; color: #f9fafb; min-height: 100vh; }
        
        header { background: rgba(17, 24, 39, 0.9); backdrop-filter: blur(12px); border-bottom: 1px solid rgba(255, 255, 255, 0.08); padding: 14px 20px; display: flex; justify-content: space-between; align-items: center; position: sticky; top: 0; z-index: 50; }
        .logo { font-size: 22px; font-weight: 900; color: #fbbf24; }
        .nav-links { display: flex; gap: 8px; align-items: center; }
        .owner-tag { background: #000; color: #fbbf24; padding: 4px 12px; border-radius: 20px; font-size: 11px; font-weight: bold; border: 1px solid #fbbf24; }
        .nav-btn { background: #1f2937; color: #fff; padding: 7px 12px; border-radius: 8px; font-size: 12px; text-decoration: none; font-weight: 700; border: 1px solid #374151; }
        .nav-btn-join { background: #059669; color: #fff; border: none; }
        
        .container { max-width: 460px; margin: 15px auto; padding: 0 12px; }
        
        .user-auth-card { background: #111827; border-radius: 16px; padding: 16px; margin-bottom: 16px; border: 1px solid rgba(255,255,255,0.08); display: flex; justify-content: space-between; align-items: center; }
        .user-info { display: flex; align-items: center; gap: 10px; }
        .user-avatar { width: 40px; height: 40px; background: #fbbf24; color: #000; border-radius: 50%; font-weight: 900; display: flex; align-items: center; justify-content: center; font-size: 18px; }
        .btn-logout { background: #374151; color: #f8fafc; border: none; padding: 6px 12px; border-radius: 6px; font-size: 11px; font-weight: 700; cursor: pointer; }

        .home-map-card { background: #111827; border-radius: 18px; padding: 12px; margin-bottom: 16px; border: 1px solid rgba(255,255,255,0.08); box-shadow: 0 10px 30px rgba(0,0,0,0.5); }
        .home-map-header { font-size: 11px; font-weight: 800; color: #34d399; text-transform: uppercase; margin-bottom: 10px; display: flex; justify-content: space-between; align-items: center; }
        #homeMap { width: 100%; height: 210px; border-radius: 12px; border: 1px solid #1f2937; }

        .leaflet-tile-container img {
            filter: brightness(0.6) invert(1) contrast(3) hue-rotate(200deg) saturate(0.3) brightness(0.7);
        }

        .booking-card { background: #111827; border-radius: 20px; padding: 22px; border: 1px solid rgba(255,255,255,0.08); box-shadow: 0 10px 30px rgba(0,0,0,0.5); }
        h2 { font-size: 18px; font-weight: 800; margin-bottom: 4px; color: #f9fafb; }
        .sub-title { font-size: 12px; color: #9ca3af; margin-bottom: 18px; }

        .input-group { margin-bottom: 14px; position: relative; }
        label { font-size: 11px; font-weight: 800; color: #9ca3af; display: block; margin-bottom: 6px; text-transform: uppercase; }
        input, select { width: 100%; padding: 13px; border: 1px solid #1f2937; border-radius: 10px; font-size: 14px; outline: none; background: #030712; color: #fff; font-weight: 600; }
        input:focus, select:focus { border-color: #fbbf24; }
        
        .suggestions-list { position: absolute; top: 100%; left: 0; right: 0; background: #1f2937; border: 1px solid #374151; border-radius: 8px; z-index: 1000; max-height: 180px; overflow-y: auto; display: none; box-shadow: 0 10px 25px rgba(0,0,0,0.8); }
        .suggestion-item { padding: 10px 12px; font-size: 12px; color: #f9fafb; cursor: pointer; border-bottom: 1px solid #374151; }
        .suggestion-item:hover { background: #374151; color: #fbbf24; }

        .gps-status-badge { background: rgba(5, 150, 105, 0.15); border: 1px solid #059669; color: #34d399; padding: 8px 12px; border-radius: 8px; font-size: 11px; font-weight: bold; margin-top: 6px; display: flex; align-items: center; justify-content: space-between; }
        .btn-gps-use { background: #059669; color: white; border: none; padding: 4px 8px; border-radius: 6px; font-size: 10px; font-weight: bold; cursor: pointer; }

        .coupon-row { display: flex; gap: 8px; }
        .btn-apply-coupon { background: #7c3aed; color: #fff; border: none; padding: 0 16px; border-radius: 10px; font-weight: 800; font-size: 12px; cursor: pointer; }

        .services, .payment-options { display: flex; gap: 10px; margin-bottom: 16px; }
        .service-box, .pay-box { flex: 1; border: 1px solid #1f2937; border-radius: 12px; padding: 12px; text-align: center; cursor: pointer; font-size: 13px; font-weight: 700; background: #030712; color: #9ca3af; }
        .service-box.selected, .pay-box.selected { border-color: #fbbf24; background: rgba(251, 191, 36, 0.1); color: #fbbf24; }
        
        .upi-container { background: #030712; border: 1px solid #fbbf24; padding: 16px; border-radius: 12px; margin-bottom: 16px; text-align: center; }
        .btn-pay-app { display: block; background: #7c3aed; color: white; padding: 12px; border-radius: 8px; text-decoration: none; font-weight: 800; font-size: 13px; margin-bottom: 10px; }
        .qr-box { background: white; padding: 8px; display: inline-block; border-radius: 10px; }
        .qr-box img { width: 110px; height: 110px; display: block; }
        .upi-id-box { background: #111827; padding: 8px; border-radius: 6px; border: 1px solid #fbbf24; font-weight: 800; color: #fbbf24; font-size: 12px; margin-top: 8px; word-break: break-all; }

        .otp-badge { background: rgba(56, 189, 248, 0.1); border: 2px dashed #38bdf8; color: #38bdf8; padding: 12px; border-radius: 10px; font-size: 16px; font-weight: 900; margin: 12px 0; text-align: center; letter-spacing: 2px; }

        .fare-preview { background: #030712; padding: 14px 16px; border-radius: 12px; margin-bottom: 16px; display: flex; justify-content: space-between; align-items: center; font-weight: 800; font-size: 15px; border: 1px solid #1f2937; }
        .btn-book { width: 100%; background: #fbbf24; color: #000; border: none; padding: 16px; font-size: 16px; font-weight: 900; border-radius: 12px; cursor: pointer; box-shadow: 0 4px 20px rgba(251, 191, 36, 0.3); }
        
        .modal { display: none; position: fixed; top: 0; left: 0; width: 100%; height: 100%; background: rgba(0,0,0,0.85); justify-content: center; align-items: center; z-index: 1000; backdrop-filter: blur(8px); }
        .modal-content { background: #111827; width: 94%; max-width: 440px; padding: 22px; border-radius: 20px; text-align: center; max-height: 94vh; overflow-y: auto; border: 1px solid rgba(255,255,255,0.1); }
        .radar-box { position: relative; width: 85px; height: 85px; margin: 15px auto; background: #030712; border-radius: 50%; display: flex; align-items: center; justify-content: center; border: 2px solid #fbbf24; animation: pulse 1.5s infinite; }
        @keyframes pulse { 0% { transform: scale(0.95); box-shadow: 0 0 0 0 rgba(251,191,36, 0.7); } 70% { transform: scale(1); box-shadow: 0 0 0 18px rgba(251,191,36, 0); } 100% { transform: scale(0.95); box-shadow: 0 0 0 0 rgba(251,191,36, 0); } }
        
        .captain-box { background: rgba(6, 78, 59, 0.4); border: 1px solid #059669; padding: 14px; border-radius: 12px; margin: 12px 0; text-align: left; color: #ecfdf5; display: none; }
        #modalMap { width: 100%; height: 220px; border-radius: 12px; margin-top: 12px; display: none; border: 1px solid #1f2937; }
        
        .action-grid { display: flex; gap: 8px; margin-top: 12px; }
        .btn-act { flex: 1; padding: 10px; border-radius: 8px; font-weight: 800; font-size: 12px; text-decoration: none; color: white; display: inline-flex; align-items: center; justify-content: center; gap: 6px; }
        .btn-call-ride { background: #0284c7; }
        .btn-wa-ride { background: #25D366; }
        .btn-chat-ride { background: #7c3aed; }

        .chat-box { background: #030712; border: 1px solid #374151; border-radius: 12px; padding: 12px; margin-top: 12px; text-align: left; display: none; }
        .chat-messages { height: 100px; overflow-y: auto; font-size: 12px; margin-bottom: 8px; display: flex; flex-direction: column; gap: 6px; padding-right: 4px; }
        .chat-msg { padding: 6px 10px; border-radius: 8px; max-width: 80%; }
        .chat-msg.customer { background: #1d4ed8; color: white; align-self: flex-end; }
        .chat-msg.captain { background: #374151; color: #f9fafb; align-self: flex-start; }
        .chat-input-row { display: flex; gap: 6px; }
        .chat-input-row input { padding: 8px; font-size: 12px; }
        .chat-send-btn { background: #fbbf24; color: black; border: none; padding: 0 12px; border-radius: 8px; font-weight: bold; cursor: pointer; }

        .rating-box { background: #030712; border: 1px solid #fbbf24; padding: 16px; border-radius: 12px; margin-top: 14px; text-align: center; display: none; }
        .star-rating { font-size: 26px; cursor: pointer; margin: 10px 0; color: #4b5563; }
        .star-rating span.active { color: #fbbf24; }

        .neon-radar-marker { background: #030712; border: 2px solid #fbbf24; width: 36px; height: 36px; border-radius: 50%; display: flex; align-items: center; justify-content: center; box-shadow: 0 0 15px rgba(251, 191, 36, 0.8); position: relative; }
        .neon-radar-marker::after { content: ''; width: 10px; height: 10px; background: #fbbf24; border-radius: 50%; box-shadow: 0 0 8px #fbbf24; }
        .user-live-gps-marker { background: #0284c7; border: 3px solid #ffffff; width: 22px; height: 22px; border-radius: 50%; box-shadow: 0 0 15px #38bdf8; }
    </style>
</head>
<body>
    <header>
        <div class="logo">⚡ RIDER</div>
        <div class="nav-links">
            <a href="/join-captain" class="nav-btn nav-btn-join">➕ Join Captain</a>
            <a href="/captain" class="nav-btn">🛵 Captain Login</a>
            <a href="/admin" class="nav-btn" style="background:#fbbf24; color:#000;">📊 Admin Login</a>
            <div class="owner-tag">{{ owner }}</div>
        </div>
    </header>
    <div class="container">
        
        <div id="profileBar" class="user-auth-card" style="display:none;">
            <div class="user-info">
                <div class="user-avatar" id="userInitial">U</div>
                <div>
                    <div id="userNameDisp" style="font-weight:800; font-size:14px; color:#fff;">User</div>
                    <div id="userPhoneDisp" style="font-size:11px; color:#9ca3af;">+91 0000000000</div>
                </div>
            </div>
            <button class="btn-logout" onclick="logoutUser()">Logout 🚪</button>
        </div>

        <div class="home-map-card">
            <div class="home-map-header">
                <span id="mapCityTitle">📍 Live City GPS Zone</span>
                <span style="background: rgba(5, 150, 105, 0.2); color: #34d399; padding: 3px 10px; border-radius: 12px; border: 1px solid #059669;">Fleet Live</span>
            </div>
            <div id="homeMap"></div>
        </div>

        <div class="booking-card">
            <h2>Book Ride or Parcel</h2>
            <div class="sub-title">📍 Dynamic Location & Instant GPS Engine</div>
            
            <form id="rideForm" onsubmit="event.preventDefault(); startBookingProcess();">
                
                <div id="authFieldsSection">
                    <div class="input-group">
                        <label>Your Name</label>
                        <input type="text" id="name" placeholder="Enter full name">
                    </div>
                    <div class="input-group">
                        <label>Phone Number (Instant Verification)</label>
                        <input type="tel" id="phone" placeholder="Enter 10-digit mobile number" maxlength="10">
                    </div>
                    <button type="button" id="btnSendOtp" class="btn-book" style="margin-bottom: 14px; background: #38bdf8; color:#000;" onclick="sendLoginOtp()">📲 Send Verification OTP</button>
                    
                    <div id="otpInputSection" style="display:none;" class="input-group">
                        <div class="otp-badge" id="otpNoticeBox">📩 Enter 4-Digit Code Below</div>
                        <label>Enter SMS OTP</label>
                        <input type="text" id="enteredOtp" placeholder="Enter 4-digit OTP" maxlength="4">
                        <button type="button" class="btn-book" style="margin-top: 10px; background: #10b981; color:#fff;" onclick="verifyAndSaveLogin()">Verify Mobile & Continue</button>
                    </div>
                </div>

                <div id="bookingSection" style="display:none;">
                    <div class="input-group">
                        <label>Pickup Location (Type or Auto-Detect)</label>
                        <input type="text" id="pickupCustom" placeholder="Search pickup (e.g. Noida, Chatro, GNIOT)" oninput="searchLocation(this.value, 'pickupSuggestions', 'pickup')" autocomplete="off" required>
                        <div id="pickupSuggestions" class="suggestions-list"></div>
                        <div class="gps-status-badge">
                            <span id="gpsStatusText">📍 Type location or use GPS</span>
                            <button type="button" class="btn-gps-use" onclick="fetchCustomerGPS()">Use Live GPS</button>
                        </div>
                        <input type="hidden" id="custLat" value="28.5355">
                        <input type="hidden" id="custLng" value="77.3910">
                    </div>

                    <div class="input-group">
                        <label>Destination Drop Location</label>
                        <input type="text" id="dropCustom" placeholder="Search destination (e.g. GNIOT, KCC, Sector 62)" oninput="searchLocation(this.value, 'dropSuggestions', 'drop')" autocomplete="off" required>
                        <div id="dropSuggestions" class="suggestions-list"></div>
                        <input type="hidden" id="dropLat" value="28.4700">
                        <input type="hidden" id="dropLng" value="77.4800">
                    </div>

                    <label>Select Service</label>
                    <div class="services">
                        <div class="service-box selected" id="box-bike" onclick="selectService('Bike Ride')">🛵 Bike Ride<br><span id="bikeRateText" style="font-size:11px; color:#9ca3af;">₹40 Base</span></div>
                        <div class="service-box" id="box-parcel" onclick="selectService('Parcel')">📦 Parcel<br><span id="parcelRateText" style="font-size:11px; color:#9ca3af;">₹60 Base</span></div>
                    </div>

                    <div class="input-group">
                        <label>Apply Coupon Code (Try: RIDER50, SONU20)</label>
                        <div class="coupon-row">
                            <input type="text" id="couponInput" placeholder="Enter coupon code" style="text-transform: uppercase;">
                            <button type="button" class="btn-apply-coupon" onclick="applyCouponCode()">Apply</button>
                        </div>
                        <div id="couponStatusText" style="font-size: 11px; margin-top: 4px; color: #34d399; font-weight: bold;"></div>
                    </div>
                    
                    <label>Payment Method</label>
                    <div class="payment-options">
                        <div class="pay-box selected" id="pay-cash" onclick="selectPayment('Cash')">💵 Cash</div>
                        <div class="pay-box" id="pay-online" onclick="selectPayment('Online / UPI')">📱 UPI / QR</div>
                    </div>

                    <div class="fare-preview">
                        <span>Estimated Fare & Distance:</span>
                        <span id="displayFare" style="color: #fbbf24; font-size: 18px;">₹40</span>
                    </div>

                    <div class="upi-container" id="upiBox">
                        <div style="font-weight: bold; color: #fbbf24; margin-bottom: 8px; font-size: 13px;">📱 Scan QR or Pay via UPI App</div>
                        <a href="#" id="upiPayLink" class="btn-pay-app" target="_blank">⚡ Pay Exact <span id="btnFareText">₹40</span> via UPI App</a>
                        <div class="qr-box">
                            <img id="upiQrImg" src="" alt="UPI QR Code">
                        </div>
                        <div class="upi-id-box">{{ upi_id }}</div>
                    </div>

                    <button type="submit" class="btn-book">Confirm Ride Booking</button>
                </div>
            </form>
        </div>
    </div>

    <div class="modal" id="trackingModal">
        <div class="modal-content">
            <div id="searchingState">
                <div class="radar-box"><div style="font-size: 30px;">⚡</div></div>
                <h3 id="searchStatusText" style="font-size: 15px; color: #f9fafb;">Looking for available Captains...</h3>
                <p style="font-size: 11px; color: #9ca3af; margin-top: 6px;">Please wait while a Captain accepts your ride request.</p>
            </div>
            
            <div class="captain-box" id="captainCard">
                <div id="statusBadgeHeader" style="color: #fbbf24; font-size: 13px; font-weight: bold; margin-bottom: 8px;">🟡 CAPTAIN ACCEPTED RIDE</div>
                <h3 style="color: #34d399; margin-bottom: 6px; font-size: 15px;">🎉 Captain Assigned!</h3>
                <div id="captainDetailsHtml" style="font-size: 13px; line-height: 1.6;"></div>
                <div class="otp-badge" id="otpDisplayBox"></div>
                
                <div style="background: #030712; border: 1px solid #fbbf24; padding: 12px; border-radius: 12px; margin: 12px 0; text-align: center;">
                    <div style="font-weight: bold; color: #fbbf24; margin-bottom: 6px; font-size: 12px;">📱 Captain / Payment QR Code</div>
                    <a href="#" id="modalUpiPayLink" class="btn-pay-app" target="_blank" style="margin-bottom:8px; padding:8px;">⚡ Pay via UPI App</a>
                    <div class="qr-box" style="padding:4px;">
                        <img id="modalUpiQrImg" src="" alt="UPI QR Code" style="width:90px; height:90px;">
                    </div>
                    <div style="font-size: 11px; color: #38bdf8; margin-top: 4px; font-weight: bold;">UPI ID: {{ upi_id }}</div>
                </div>

                <div class="action-grid">
                    <a id="callCaptainBtn" href="#" class="btn-act btn-call-ride">📞 Call</a>
                    <a id="whatsappCaptainBtn" href="#" target="_blank" class="btn-act btn-wa-ride">💬 WhatsApp</a>
                    <button onclick="toggleChatBox()" class="btn-act btn-chat-ride">💬 In-App Chat</button>
                </div>

                <div class="chat-box" id="inAppChatBox">
                    <div class="chat-messages" id="chatMessagesArea">
                        <div class="chat-msg captain">Hello ji! Main aapki location ki taraf aa raha hoon.</div>
                    </div>
                    <div class="chat-input-row">
                        <input type="text" id="chatMsgInput" placeholder="Type message to captain...">
                        <button class="chat-send-btn" onclick="sendChatMessage()">Send</button>
                    </div>
                </div>
            </div>
            
            <div id="modalMap"></div>

            <div class="rating-box" id="customerRatingSection">
                <h4 style="color: #fbbf24; font-size: 15px;">⭐ Rate Your Captain Experience</h4>
                <div class="star-rating" id="starContainer">
                    <span onclick="setRating(1)">★</span>
                    <span onclick="setRating(2)">★</span>
                    <span onclick="setRating(3)">★</span>
                    <span onclick="setRating(4)">★</span>
                    <span onclick="setRating(5)" class="active">★</span>
                </div>
                <button onclick="submitCustomerRating()" class="btn-book" style="padding: 10px; font-size: 13px;">Submit Rating & Review</button>
            </div>
            
            <br>
            <button id="resetBtn" onclick="location.reload()" class="btn-book" style="display:none; margin-top: 14px;">Book Another Ride 🛵</button>
        </div>
    </div>

    <script src="https://unpkg.com/leaflet@1.9.4/dist/leaflet.js"></script>
    <script>
        let selectedServiceType = 'Bike Ride';
        let selectedPaymentMode = 'Cash';
        let baseCalculatedFare = 40;
        let appliedDiscount = 0;
        let appliedCouponName = 'None';
        const ownerUpi = "{{ upi_id }}";
        let modalMapObj = null;
        let captainTrackingMarker = null;
        let hMap = null;
        let userGpsMarker = null;
        let currentBookingIndex = null;
        let statusPollInterval = null;
        let selectedRatingValue = 5;

        function checkUserProfile() {
            const savedName = localStorage.getItem('rider_user_name');
            const savedPhone = localStorage.getItem('rider_user_phone');
            
            if (savedName && savedPhone) {
                document.getElementById('authFieldsSection').style.display = 'none';
                document.getElementById('bookingSection').style.display = 'block';
                document.getElementById('profileBar').style.display = 'flex';
                document.getElementById('userNameDisp').innerText = savedName;
                document.getElementById('userPhoneDisp').innerText = "+91 " + savedPhone;
                document.getElementById('userInitial').innerText = savedName.charAt(0).toUpperCase();
            } else {
                document.getElementById('authFieldsSection').style.display = 'block';
                document.getElementById('bookingSection').style.display = 'none';
                document.getElementById('profileBar').style.display = 'none';
            }
        }

        function sendLoginOtp() {
            const name = document.getElementById('name').value.trim();
            const phone = document.getElementById('phone').value.trim();
            const phoneRegex = /^[6-9][0-9]{9}$/;

            if (!name) { alert("Please enter your name!"); return; }
            if (!phoneRegex.test(phone)) { alert("Please enter a valid 10-digit Indian mobile number!"); return; }

            fetch('/api/send-sms-otp', {
                method: 'POST',
                headers: {'Content-Type': 'application/json'},
                body: JSON.stringify({ name, phone })
            })
            .then(res => res.json())
            .then(data => {
                document.getElementById('otpInputSection').style.display = 'block';
                document.getElementById('btnSendOtp').style.display = 'none';
                if(data.otp) {
                    alert("🔑 Your Verification OTP is: " + data.otp);
                    document.getElementById('otpNoticeBox').innerHTML = `🔑 ACTIVE OTP: <b style="color:#fff; font-size:18px;">${data.otp}</b>`;
                }
            });
        }

        function verifyAndSaveLogin() {
            const phone = document.getElementById('phone').value.trim();
            const entered = document.getElementById('enteredOtp').value.trim();

            fetch('/api/verify-sms-otp', {
                method: 'POST',
                headers: {'Content-Type': 'application/json'},
                body: JSON.stringify({ phone, otp: entered })
            })
            .then(res => res.json())
            .then(data => {
                if (data.status === "success") {
                    const name = document.getElementById('name').value.trim();
                    localStorage.setItem('rider_user_name', name);
                    localStorage.setItem('rider_user_phone', phone);
                    alert("✅ Mobile Verified Successfully!");
                    checkUserProfile();
                } else {
                    alert("❌ Wrong OTP! Please enter correct 4-digit code.");
                }
            });
        }

        function logoutUser() {
            if(confirm("Are you sure you want to log out?")) {
                localStorage.removeItem('rider_user_name');
                localStorage.removeItem('rider_user_phone');
                location.reload();
            }
        }

        function applyCouponCode() {
            const code = document.getElementById('couponInput').value.trim().toUpperCase();
            fetch('/api/apply-coupon', {
                method: 'POST',
                headers: {'Content-Type': 'application/json'},
                body: JSON.stringify({ coupon: code, base_fare: baseCalculatedFare })
            })
            .then(res => res.json())
            .then(data => {
                if(data.status === "success") {
                    appliedDiscount = data.discount;
                    appliedCouponName = code;
                    document.getElementById('couponStatusText').innerText = `🎉 Coupon '${code}' applied! Saved ₹${data.discount}`;
                    alert(`🎉 Coupon Applied Successfully! You saved ₹${data.discount}`);
                    calculateFare();
                } else {
                    alert("❌ Invalid or Expired Coupon Code!");
                }
            });
        }

        function searchLocation(query, listId, type) {
            if(query.length < 1) {
                document.getElementById(listId).style.display = 'none';
                return;
            }
            fetch(`/api/proxy-search?q=${encodeURIComponent(query)}`)
            .then(res => res.json())
            .then(data => {
                const listDiv = document.getElementById(listId);
                listDiv.innerHTML = '';
                if(data && data.length > 0) {
                    listDiv.style.display = 'block';
                    data.forEach(item => {
                        const div = document.createElement('div');
                        div.className = 'suggestion-item';
                        div.innerText = item.display_name;
                        div.onclick = function() {
                            if(type === 'pickup') {
                                document.getElementById('pickupCustom').value = item.display_name;
                                document.getElementById('custLat').value = item.lat;
                                document.getElementById('custLng').value = item.lon;
                            } else {
                                document.getElementById('dropCustom').value = item.display_name;
                                document.getElementById('dropLat').value = item.lat;
                                document.getElementById('dropLng').value = item.lon;
                            }
                            listDiv.style.display = 'none';
                            updateDynamicFareAndDistance();
                        };
                        listDiv.appendChild(div);
                    });
                } else {
                    listDiv.style.display = 'none';
                }
            }).catch(err => { console.log("Search proxy error:", err); });
        }

        function calculateDistance(lat1, lon1, lat2, lon2) {
            const R = 6371;
            const dLat = (lat2 - lat1) * Math.PI / 180;
            const dLon = (lon2 - lon1) * Math.PI / 180;
            const a = Math.sin(dLat/2) * Math.sin(dLat/2) +
                      Math.cos(lat1 * Math.PI / 180) * Math.cos(lat2 * Math.PI / 180) *
                      Math.sin(dLon/2) * Math.sin(dLon/2);
            const c = 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1-a));
            return R * c;
        }

        function updateDynamicFareAndDistance() {
            const lat1 = parseFloat(document.getElementById('custLat').value);
            const lon1 = parseFloat(document.getElementById('custLng').value);
            const lat2 = parseFloat(document.getElementById('dropLat').value);
            const lon2 = parseFloat(document.getElementById('dropLng').value);

            const distKm = calculateDistance(lat1, lon1, lat2, lon2);
            
            let baseRate = (selectedServiceType === 'Parcel') ? 60 : 40;
            let calculated = Math.round(baseRate + (distKm * 12));
            if(calculated < baseRate) calculated = baseRate;
            
            baseCalculatedFare = calculated;
            calculateFare();
        }

        function fetchCustomerGPS() {
            if (navigator.geolocation) {
                document.getElementById('gpsStatusText').innerText = "📍 Fetching Live GPS Coordinates...";
                navigator.geolocation.getCurrentPosition(position => {
                    const lat = position.coords.latitude;
                    const lng = position.coords.longitude;
                    document.getElementById('custLat').value = lat;
                    document.getElementById('custLng').value = lng;
                    
                    fetch(`/api/proxy-reverse?lat=${lat}&lon=${lng}`)
                    .then(res => res.json())
                    .then(data => {
                        const displayName = data.display_name ? data.display_name.split(',').slice(0, 3).join(',') : `GPS: ${lat.toFixed(3)}, ${lng.toFixed(3)}`;
                        document.getElementById('pickupCustom').value = displayName;
                        document.getElementById('gpsStatusText').innerText = `📍 Live GPS: ${displayName}`;
                    });

                    if (hMap) {
                        hMap.setView([lat, lng], 15);
                        if (userGpsMarker) {
                            userGpsMarker.setLatLng([lat, lng]);
                        } else {
                            const userDotIcon = L.divIcon({ html: `<div class="user-live-gps-marker"></div>`, className: '', iconSize: [22, 22], iconAnchor: [11, 11] });
                            userGpsMarker = L.marker([lat, lng], { icon: userDotIcon }).addTo(hMap).bindPopup("<b>Your Live GPS Location</b>").openPopup();
                        }
                    }
                    updateDynamicFareAndDistance();
                }, error => {
                    document.getElementById('gpsStatusText').innerText = "⚠️ Please enable GPS location access in browser";
                }, { enableHighAccuracy: true });
            }
        }

        function initHomeMap() {
            const defaultLat = 28.5355;
            const defaultLng = 77.3910;
            hMap = L.map('homeMap', { zoomControl: false }).setView([defaultLat, defaultLng], 13);
            
            L.tileLayer('https://tile.openstreetmap.org/{z}/{x}/{y}.png', {
                maxZoom: 19,
                attribution: '© OpenStreetMap'
            }).addTo(hMap);

            fetchCustomerGPS();
        }

        function selectService(serviceName) {
            selectedServiceType = serviceName;
            document.querySelectorAll('.services .service-box').forEach(el => el.classList.remove('selected'));
            if(serviceName === 'Bike Ride') document.getElementById('box-bike').classList.add('selected');
            else document.getElementById('box-parcel').classList.add('selected');
            updateDynamicFareAndDistance();
        }

        function selectPayment(mode) {
            selectedPaymentMode = mode;
            document.querySelectorAll('.payment-options .pay-box').forEach(el => el.classList.remove('selected'));
            
            if(mode === 'Cash') {
                document.getElementById('pay-cash').classList.add('selected');
            } else {
                document.getElementById('pay-online').classList.add('selected');
            }
            updateUpiDetails(Math.max(10, baseCalculatedFare - appliedDiscount));
        }

        function calculateFare() {
            let finalFare = Math.max(10, baseCalculatedFare - appliedDiscount);
            const lat1 = parseFloat(document.getElementById('custLat').value);
            const lon1 = parseFloat(document.getElementById('custLng').value);
            const lat2 = parseFloat(document.getElementById('dropLat').value);
            const lon2 = parseFloat(document.getElementById('dropLng').value);
            const distKm = calculateDistance(lat1, lon1, lat2, lon2).toFixed(1);

            document.getElementById('displayFare').innerText = `₹${finalFare} (${distKm} km)`;
            document.getElementById('btnFareText').innerText = '₹' + finalFare;
            updateUpiDetails(finalFare);
        }

        function updateUpiDetails(amount) {
            const upiUrl = `upi://pay?pa=${ownerUpi}&pn=Sonu%20Kumar&am=${amount}&cu=INR&tn=Rider%20Booking`;
            document.getElementById('upiPayLink').setAttribute('href', upiUrl);
            document.getElementById('modalUpiPayLink').setAttribute('href', upiUrl);
            
            const qrApiUrl = `https://api.qrserver.com/v1/create-qr-code/?size=110x110&data=${encodeURIComponent(upiUrl)}`;
            document.getElementById('upiQrImg').setAttribute('src', qrApiUrl);
            document.getElementById('modalUpiQrImg').setAttribute('src', qrApiUrl);
        }

        function initTrackingMap(lat, lng) {
            document.getElementById('modalMap').style.display = 'block';
            if (!modalMapObj) {
                modalMapObj = L.map('modalMap').setView([lat, lng], 14);
                L.tileLayer('https://tile.openstreetmap.org/{z}/{x}/{y}.png', { maxZoom: 19 }).addTo(modalMapObj);
                const customBikeIcon = L.divIcon({ html: '<div class="neon-radar-marker"></div>', className: '', iconSize: [36, 36], iconAnchor: [18, 18] });
                captainTrackingMarker = L.marker([lat, lng], {icon: customBikeIcon}).addTo(modalMapObj).bindPopup("<b>Captain is coming to your location!</b>").openPopup();
            } else {
                modalMapObj.setView([lat, lng], 14);
                captainTrackingMarker.setLatLng([lat, lng]);
            }
            setTimeout(() => { modalMapObj.invalidateSize(); }, 250);
        }

        function toggleChatBox() {
            const box = document.getElementById('inAppChatBox');
            box.style.display = box.style.display === 'block' ? 'none' : 'block';
        }

        function sendChatMessage() {
            const input = document.getElementById('chatMsgInput');
            const txt = input.value.trim();
            if(!txt) return;

            const area = document.getElementById('chatMessagesArea');
            area.innerHTML += `<div class="chat-msg customer">${txt}</div>`;
            input.value = '';
            area.scrollTop = area.scrollHeight;

            setTimeout(() => {
                area.innerHTML += `<div class="chat-msg captain">Theek hai ji, main location par pahunch raha hoon!</div>`;
                area.scrollTop = area.scrollHeight;
            }, 1200);
        }

        function setRating(val) {
            selectedRatingValue = val;
            const stars = document.querySelectorAll('#starContainer span');
            stars.forEach((s, idx) => {
                if (idx < val) s.classList.add('active');
                else s.classList.remove('active');
            });
        }

        function submitCustomerRating() {
            if (currentBookingIndex === null) return;

            fetch('/api/submit-rating', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ index: currentBookingIndex, rating: selectedRatingValue + " ⭐" })
            })
            .then(res => res.json())
            .then(data => {
                alert("🎉 Thank you for your rating!");
                document.getElementById('customerRatingSection').style.display = 'none';
                document.getElementById('resetBtn').style.display = 'block';
            });
        }

        function pollRideStatus() {
            if (currentBookingIndex === null) return;

            fetch(`/api/booking-status/${currentBookingIndex}`)
            .then(res => res.json())
            .then(data => {
                if (data.status === "Accepted") {
                    document.getElementById('searchingState').style.display = 'none';
                    document.getElementById('captainCard').style.display = 'block';
                    document.getElementById('statusBadgeHeader').innerHTML = `🟡 RIDE ACCEPTED BY CAPTAIN`;
                    document.getElementById('statusBadgeHeader').style.color = '#fbbf24';
                    
                    document.getElementById('captainDetailsHtml').innerHTML = `
                        <strong>👤 ${data.captain.name}</strong> (${data.captain.rating})<br>
                        <span style="color: #a7f3d0;">Bike: ${data.captain.bike}</span><br>
                        <span style="color: #34d399; font-weight: bold;">Fare: ${data.fare} (${data.payment_mode})</span>
                    `;
                    document.getElementById('otpDisplayBox').innerHTML = `🔑 RIDE VERIFICATION OTP: <span style="color:#38bdf8; font-size:18px;">${data.otp}</span>`;
                    document.getElementById('callCaptainBtn').setAttribute('href', `tel:${data.captain.phone}`);
                    document.getElementById('whatsappCaptainBtn').setAttribute('href', data.whatsapp_url);

                    const lat = parseFloat(document.getElementById('custLat').value);
                    const lng = parseFloat(document.getElementById('custLng').value);
                    initTrackingMap(lat, lng);
                } else if (data.status === "Completed") {
                    document.getElementById('statusBadgeHeader').innerHTML = `🎉 RIDE COMPLETED!`;
                    document.getElementById('statusBadgeHeader').style.color = '#34d399';
                    document.getElementById('customerRatingSection').style.display = 'block';
                    clearInterval(statusPollInterval);
                }
            });
        }

        function startBookingProcess() {
            const name = localStorage.getItem('rider_user_name') || "Customer";
            const phone = localStorage.getItem('rider_user_phone') || "9999999999";
            const pickup = document.getElementById('pickupCustom').value;
            const drop = document.getElementById('dropCustom').value;
            const lat = document.getElementById('custLat').value;
            const lng = document.getElementById('custLng').value;
            const finalFareText = document.getElementById('displayFare').innerText;
            
            if(!pickup || !drop) { alert("Please enter both pickup and destination drop locations!"); return; }

            document.getElementById('trackingModal').style.display = 'flex';
            document.getElementById('searchingState').style.display = 'block';
            document.getElementById('captainCard').style.display = 'none';
            document.getElementById('modalMap').style.display = 'none';
            document.getElementById('customerRatingSection').style.display = 'none';
            document.getElementById('resetBtn').style.display = 'none';

            fetch('/api/book', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ 
                    name, phone, pickup, drop, 
                    service: selectedServiceType, 
                    fare: finalFareText, 
                    payment_mode: selectedPaymentMode, 
                    lat, lng, 
                    coupon: appliedCouponName 
                })
            })
            .then(response => response.json())
            .then(data => {
                if(data.status === "success") {
                    currentBookingIndex = data.booking_index;
                    statusPollInterval = setInterval(pollRideStatus, 2500);
                }
            });
        }
        
        window.onload = function() {
            checkUserProfile();
            calculateFare();
            initHomeMap();
        };
    </script>
</body>
</html>
"""

ADMIN_TEMPLATE = """
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Admin Dashboard - Rider</title>
    <meta http-equiv="refresh" content="10">
    <style>
        * { box-sizing: border-box; margin: 0; padding: 0; font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif; }
        body { background-color: #030712; color: #f9fafb; padding: 20px; }
        .header { display: flex; justify-content: space-between; align-items: center; background: #111827; border: 1px solid rgba(255,255,255,0.08); color: #fbbf24; padding: 15px 25px; border-radius: 12px; margin-bottom: 20px; flex-wrap: wrap; gap: 10px; }
        .btn-home { background: #fbbf24; color: #000; padding: 8px 15px; border-radius: 6px; text-decoration: none; font-weight: bold; font-size: 13px; }
        .card { background: #111827; border: 1px solid rgba(255,255,255,0.08); border-radius: 12px; padding: 20px; overflow-x: auto; margin-bottom: 20px; }
        .stats-grid { display: flex; gap: 15px; flex-wrap: wrap; margin-bottom: 20px; }
        .stat-card { background: #1f2937; padding: 16px; border-radius: 10px; flex: 1; min-width: 180px; border: 1px solid rgba(255,255,255,0.05); }
        h2 { color: #f9fafb; font-size: 18px; margin-bottom: 10px; }
        table { width: 100%; border-collapse: collapse; margin-top: 10px; }
        th, td { padding: 12px 15px; text-align: left; border-bottom: 1px solid #1f2937; font-size: 13px; color: #d1d5db; }
        th { background-color: #1f2937; color: #fbbf24; text-transform: uppercase; font-size: 11px; }
        .btn-status-approved { padding: 4px 10px; border-radius: 6px; font-weight: bold; font-size: 11px; color: white; background: #059669; }
        .btn-status-pending { padding: 4px 10px; border-radius: 6px; font-weight: bold; font-size: 11px; color: white; background: #eab308; }
        .btn-approve-act { padding: 6px 12px; border-radius: 6px; font-weight: bold; font-size: 11px; color: white; background: #0284c7; text-decoration: none; display: inline-block; }
    </style>
</head>
<body>
    <div class="header">
        <h1>⚡ RIDER - Admin Control Console</h1>
        <div>
            <a href="/logout" class="btn-home" style="background:#f43f5e; color:white;">Logout 🚪</a>
            <a href="/" class="btn-home">⬅ Back to Booking</a>
        </div>
    </div>

    <div class="stats-grid">
        <div class="stat-card">
            <div style="font-size: 11px; color: #9ca3af; font-weight: bold;">TOTAL GROSS BOOKINGS</div>
            <div style="font-size: 22px; font-weight: 900; color: #fff; margin-top: 4px;">₹{{ total_gross_earnings }}</div>
        </div>
        <div class="stat-card" style="border-left: 4px solid #fbbf24;">
            <div style="font-size: 11px; color: #fbbf24; font-weight: bold;">ADMIN COMMISSION (20%)</div>
            <div style="font-size: 22px; font-weight: 900; color: #fbbf24; margin-top: 4px;">₹{{ total_admin_commission }}</div>
        </div>
        <div class="stat-card" style="border-left: 4px solid #34d399;">
            <div style="font-size: 11px; color: #34d399; font-weight: bold;">CAPTAINS PAYOUT POOL (80%)</div>
            <div style="font-size: 22px; font-weight: 900; color: #34d399; margin-top: 4px;">₹{{ total_captain_payout }}</div>
        </div>
    </div>

    <div class="card">
        <h2>🛵 Registered Captains Verification & Approval Panel</h2>
        <table>
            <thead>
                <tr>
                    <th>Captain Name</th>
                    <th>Phone</th>
                    <th>Bike Model</th>
                    <th>Vehicle Number</th>
                    <th>Licence / ID</th>
                    <th>Status</th>
                    <th>Action</th>
                </tr>
            </thead>
            <tbody>
                {% for idx, row in captains_with_index %}
                <tr>
                    <td><b>{{ row[0] if row|length > 0 else '' }}</b></td>
                    <td>{{ row[1] if row|length > 1 else '' }}</td>
                    <td>{{ row[2] if row|length > 2 else '' }}</td>
                    <td><b style="color:#fbbf24;">{{ row[3] if row|length > 3 else '' }}</b></td>
                    <td>{{ row[4] if row|length > 4 else '' }}</td>
                    <td>
                        {% if row|length > 5 and row[5] == 'Approved' %}
                        <span class="btn-status-approved">Approved ✅</span>
                        {% else %}
                        <span class="btn-status-pending">Pending ⏳</span>
                        {% endif %}
                    </td>
                    <td>
                        {% if row|length > 5 and row[5] != 'Approved' %}
                        <a href="/admin/approve-captain/{{ idx }}" class="btn-approve-act">Click to Approve ✅</a>
                        {% else %}
                        <span style="color:#34d399; font-weight:bold;">Active</span>
                        {% endif %}
                    </td>
                </tr>
                {% endfor %}
            </tbody>
        </table>
    </div>

    <div class="card">
        <h2>📊 Live Ride Bookings & Commission Breakdown</h2>
        <table>
            <thead>
                <tr>
                    <th>Customer</th>
                    <th>Pickup $\rightarrow$ Drop</th>
                    <th>Fare Paid</th>
                    <th>Admin (20%)</th>
                    <th>Rider (80%)</th>
                    <th>Captain</th>
                    <th>OTP</th>
                    <th>Rating</th>
                    <th>Status</th>
                </tr>
            </thead>
            <tbody>
                {% for fare_num, admin_c, captain_p, row in bookings_data %}
                <tr>
                    <td><b>{{ row[0] if row|length > 0 else '' }}</b><br><span style="font-size:11px; color:#9ca3af;">{{ row[1] if row|length > 1 else '' }}</span></td>
                    <td>{{ row[2] if row|length > 2 else '' }} $\rightarrow$ {{ row[3] if row|length > 3 else '' }}</td>
                    <td><b style="color:#fff;">₹{{ fare_num }}</b></td>
                    <td><b style="color:#fbbf24;">₹{{ admin_c }}</b></td>
                    <td><b style="color:#34d399;">₹{{ captain_p }}</b></td>
                    <td><b>{{ row[7] if row|length > 7 else 'Searching...' }}</b></td>
                    <td><b style="color: #38bdf8;">{{ row[9] if row|length > 9 else '----' }}</b></td>
                    <td><b style="color: #fbbf24;">{{ row[12] if row|length > 12 else 'Not Rated' }}</b></td>
                    <td><span style="font-weight: bold; {% if row|length > 8 and row[8] == 'Accepted' %}color: #fbbf24;{% elif row|length > 8 and row[8] == 'Completed' %}color: #34d399;{% else %}color: #f43f5e;{% endif %}">{{ row[8] if row|length > 8 else 'Pending' }}</span></td>
                </tr>
                {% endfor %}
            </tbody>
        </table>
    </div>
</body>
</html>
"""

CAPTAIN_TEMPLATE = """
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Captain Operations Console</title>
    <meta http-equiv="refresh" content="6">
    <style>
        * { box-sizing: border-box; margin: 0; padding: 0; font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif; }
        body { background-color: #030712; color: #f9fafb; padding: 20px; }
        .header { display: flex; justify-content: space-between; align-items: center; background: #064e3b; color: #fff; padding: 15px 25px; border-radius: 12px; margin-bottom: 20px; flex-wrap: wrap; gap: 10px; }
        .btn-home { background: #fbbf24; color: #000; padding: 8px 15px; border-radius: 6px; text-decoration: none; font-weight: bold; font-size: 13px; }
        .card { background: #111827; border-radius: 12px; padding: 20px; border: 1px solid rgba(255,255,255,0.08); overflow-x: auto; margin-bottom: 20px; }
        .stats-grid { display: flex; gap: 15px; flex-wrap: wrap; margin-bottom: 20px; }
        .stat-card { background: rgba(6, 78, 59, 0.4); padding: 16px; border-radius: 10px; flex: 1; min-width: 180px; border: 1px solid #059669; }
        h2 { color: #f9fafb; font-size: 18px; margin-bottom: 12px; }
        table { width: 100%; border-collapse: collapse; margin-top: 10px; }
        th, td { padding: 12px 15px; text-align: left; border-bottom: 1px solid #1f2937; font-size: 13px; color: #d1d5db; }
        th { background-color: #1f2937; color: #34d399; text-transform: uppercase; font-size: 11px; }
        .otp-input { width: 100px; padding: 6px; border: 1px solid #059669; border-radius: 6px; font-size: 13px; font-weight: bold; background: #030712; color: white; outline: none; }
        .btn-action { padding: 6px 12px; border-radius: 6px; font-weight: bold; font-size: 12px; border: none; cursor: pointer; color: white; text-decoration: none; }
        .btn-accept { background: #059669; }
        .btn-complete { background: #0284c7; }
    </style>
</head>
<body>
    <div class="header">
        <h1>🛵 Welcome Captain: {{ captain_name }}</h1>
        <div>
            <a href="/logout" class="btn-home" style="background:#f43f5e; color:white;">Logout 🚪</a>
            <a href="/" class="btn-home">⬅ Back to Booking</a>
        </div>
    </div>

    <div class="stats-grid">
        <div class="stat-card">
            <div style="font-size: 11px; color: #a7f3d0; font-weight: bold;">TOTAL COMPLETED RIDES</div>
            <div style="font-size: 22px; font-weight: 900; color: #fff; margin-top: 4px;">{{ completed_count }} Rides</div>
        </div>
        <div class="stat-card">
            <div style="font-size: 11px; color: #34d399; font-weight: bold;">YOUR NET EARNINGS (80% SHARE)</div>
            <div style="font-size: 22px; font-weight: 900; color: #34d399; margin-top: 4px;">₹{{ total_captain_net }}</div>
        </div>
        <div class="stat-card">
            <div style="font-size: 11px; color: #fbbf24; font-weight: bold;">PLATFORM COMMISSION PAID (20%)</div>
            <div style="font-size: 22px; font-weight: 900; color: #fbbf24; margin-top: 4px;">₹{{ total_admin_cut }}</div>
        </div>
    </div>

    <div class="card">
        <h2>Live Ride Requests & Tasks</h2>
        <table>
            <thead>
                <tr>
                    <th>Customer</th>
                    <th>Pickup $\rightarrow$ Drop</th>
                    <th>Total Fare</th>
                    <th>Your Share (80%)</th>
                    <th>Admin Fee (20%)</th>
                    <th>Status</th>
                    <th>Action</th>
                </tr>
            </thead>
            <tbody>
                {% for fare_num, admin_c, captain_p, index, row in bookings_data %}
                <tr>
                    <td><b>{{ row[0] }}</b><br><span style="color:#9ca3af; font-size:11px;">📱 {{ row[1] }}</span></td>
                    <td><b>📍 {{ row[2] }}</b> $\rightarrow$ 🎯 {{ row[3] }}</td>
                    <td><b>₹{{ fare_num }}</b></td>
                    <td><b style="color:#34d399;">₹{{ captain_p }}</b></td>
                    <td><span style="color:#fbbf24;">₹{{ admin_c }}</span></td>
                    <td>
                        <span style="font-weight: bold; {% if row[8] == 'Accepted' %}color: #fbbf24;{% elif row[8] == 'Completed' %}color: #34d399;{% else %}color: #f43f5e;{% endif %}">
                            {{ row[8] }}
                        </span>
                    </td>
                    <td>
                        {% if row[8] == 'Pending' %}
                        <a href="/captain/accept/{{ index }}" class="btn-action btn-accept">Accept Ride ⚡</a>
                        {% elif row[8] == 'Accepted' %}
                        <input type="text" id="otp_{{ index }}" class="otp-input" placeholder="Enter OTP">
                        <button onclick="completeRide({{ index }}, '{{ row[9] }}')" class="btn-action btn-complete">Complete</button>
                        {% else %}
                        <b style="color:#34d399;">Completed ✅</b>
                        {% endif %}
                    </td>
                </tr>
                {% endfor %}
            </tbody>
        </table>
    </div>

    <script>
        function completeRide(index, correctOtp) {
            const entered = document.getElementById('otp_' + index).value.trim();
            if (entered !== correctOtp) {
                alert("❌ Invalid Ride OTP! Ask customer for correct code.");
                return;
            }
            window.location.href = `/captain/complete/${index}?otp=${entered}`;
        }
    </script>
</body>
</html>
"""

@app.route("/")
def home():
    return render_template_string(HTML_TEMPLATE, owner=OWNER_NAME, upi_id=OWNER_UPI_ID)

@app.route("/logout")
def logout():
    session.clear()
    return redirect("/")

@app.route("/admin", methods=["GET", "POST"])
def admin_dashboard():
    error = None
    if request.method == "POST":
        pwd = request.form.get("password")
        if pwd == ADMIN_PASSWORD:
            session["is_admin"] = True
        else:
            error = "Invalid Secret Password!"

    if not session.get("is_admin"):
        return render_template_string(LOGIN_FORM_HTML, role="Admin", error=error)

    init_excel()
    wb_bookings = load_workbook(EXCEL_FILE)
    ws_bookings = wb_bookings["Bookings"]
    all_bookings = list(ws_bookings.iter_rows(values_only=True))

    bookings_data = []
    total_gross_earnings = 0
    total_admin_commission = 0
    total_captain_payout = 0

    for row in all_bookings[1:]:
        fare_raw = str(row[5]).replace("₹", "").split("(")[0].strip() if len(row) > 5 else "0"
        try:
            fare_num = int(fare_raw)
        except ValueError:
            fare_num = 0
        
        admin_commission = int(fare_num * 0.20)
        captain_payout = fare_num - admin_commission

        if len(row) > 8 and row[8] == "Completed":
            total_gross_earnings += fare_num
            total_admin_commission += admin_commission
            total_captain_payout += captain_payout

        bookings_data.append((fare_num, admin_commission, captain_payout, row))

    wb_captains = load_workbook(CAPTAINS_FILE)
    ws_captains = wb_captains["Captains"]
    all_captains = list(ws_captains.iter_rows(values_only=True))
    captains_with_index = [(i, row) for i, row in enumerate(all_captains[1:])]

    return render_template_string(ADMIN_TEMPLATE, 
                                 bookings_data=bookings_data, 
                                 total_gross_earnings=total_gross_earnings,
                                 total_admin_commission=total_admin_commission,
                                 total_captain_payout=total_captain_payout,
                                 captains_with_index=captains_with_index)

@app.route("/captain", methods=["GET", "POST"])
def captain_dashboard():
    error = None
    if request.method == "POST":
        phone = request.form.get("phone", "").strip()
        init_excel()
        wb = load_workbook(CAPTAINS_FILE)
        ws = wb["Captains"]
        captains = list(ws.iter_rows(values_only=True))
        
        valid_captain = None
        for cap in captains[1:]:
            if str(cap[1]).strip() == phone and cap[5] == "Approved":
                valid_captain = cap
                break
        
        if valid_captain:
            session["captain_phone"] = phone
            session["captain_name"] = valid_captain[0]
        else:
            error = "Captain Phone not registered OR Registration Pending Approval!"

    if not session.get("captain_phone"):
        return render_template_string(LOGIN_FORM_HTML, role="Captain", error=error)

    init_excel()
    wb = load_workbook(EXCEL_FILE)
    ws = wb["Bookings"]
    all_rows = list(ws.iter_rows(values_only=True))
    
    bookings_data = []
    completed_count = 0
    total_captain_net = 0
    total_admin_cut = 0

    for i, row in enumerate(all_rows[1:]):
        fare_raw = str(row[5]).replace("₹", "").split("(")[0].strip() if len(row) > 5 else "0"
        try:
            fare_num = int(fare_raw)
        except ValueError:
            fare_num = 0
        
        admin_commission = int(fare_num * 0.20)
        captain_payout = fare_num - admin_commission

        if len(row) > 8 and row[8] == "Completed":
            completed_count += 1
            total_captain_net += captain_payout
            total_admin_cut += admin_commission

        bookings_data.append((fare_num, admin_commission, captain_payout, i, row))

    return render_template_string(CAPTAIN_TEMPLATE, bookings_data=bookings_data, completed_count=completed_count, total_captain_net=total_captain_net, total_admin_cut=total_admin_cut, captain_name=session.get("captain_name"))

@app.route("/captain/accept/<int:index>")
def captain_accept(index):
    if not session.get("captain_phone"):
        return redirect("/captain")
    init_excel()
    wb = load_workbook(EXCEL_FILE)
    ws = wb["Bookings"]
    row_num = index + 2
    if row_num <= ws.max_row:
        ws.cell(row=row_num, column=8, value=session.get("captain_name", "Mukesh Kumar"))
        ws.cell(row=row_num, column=9, value="Accepted")
        wb.save(EXCEL_FILE)
    return redirect("/captain")

@app.route("/captain/complete/<int:index>")
def captain_complete(index):
    if not session.get("captain_phone"):
        return redirect("/captain")
    init_excel()
    wb = load_workbook(EXCEL_FILE)
    ws = wb["Bookings"]
    row_num = index + 2
    if row_num <= ws.max_row:
        ws.cell(row=row_num, column=9, value="Completed")
        wb.save(EXCEL_FILE)
    return redirect("/captain")

@app.route("/admin/approve-captain/<int:index>")
def approve_captain(index):
    if not session.get("is_admin"):
        return redirect("/admin")
    init_excel()
    wb = load_workbook(CAPTAINS_FILE)
    ws = wb["Captains"]
    row_num = index + 2
    if row_num <= ws.max_row:
        ws.cell(row=row_num, column=6, value="Approved")
        wb.save(CAPTAINS_FILE)
    return redirect("/admin")

@app.route("/join-captain")
def join_captain_page():
    return render_template_string("""
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Join as Captain - Rider</title>
    <style>
        * { box-sizing: border-box; margin: 0; padding: 0; font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif; }
        body { background-color: #030712; color: #f9fafb; padding: 20px; }
        .header { display: flex; justify-content: space-between; align-items: center; background: #111827; color: #fbbf24; padding: 15px 25px; border-radius: 12px; margin-bottom: 20px; border: 1px solid rgba(255,255,255,0.08); }
        .btn-home { background: #fbbf24; color: #000; padding: 8px 15px; border-radius: 6px; text-decoration: none; font-weight: bold; font-size: 13px; }
        .card { background: #111827; border-radius: 14px; padding: 25px; max-width: 500px; margin: 0 auto; border: 1px solid rgba(255,255,255,0.08); }
        h2 { margin-bottom: 8px; color: #f9fafb; }
        .sub { font-size: 12px; color: #9ca3af; margin-bottom: 20px; }
        .input-group { margin-bottom: 15px; }
        label { font-size: 11px; font-weight: 800; color: #9ca3af; display: block; margin-bottom: 5px; text-transform: uppercase; }
        input { width: 100%; padding: 12px; border: 1px solid #1f2937; border-radius: 8px; font-size: 14px; outline: none; background: #030712; color: #fff; }
        .btn-submit { width: 100%; background: #059669; color: white; border: none; padding: 14px; font-size: 15px; font-weight: bold; border-radius: 8px; cursor: pointer; margin-top: 10px; }
    </style>
</head>
<body>
    <div class="header">
        <h1>🛵 RIDER - Captain Partner Program</h1>
        <a href="/" class="btn-home">⬅ Back to Home</a>
    </div>
    <div class="card">
        <h2>Register as a Captain</h2>
        <div class="sub">Fill your details & vehicle documents to start earning.</div>
        <form action="/api/register-captain" method="POST">
            <div class="input-group">
                <label>Full Name</label>
                <input type="text" name="name" required placeholder="Enter full name">
            </div>
            <div class="input-group">
                <label>Mobile Number</label>
                <input type="tel" name="phone" required placeholder="10-digit mobile number" maxlength="10">
            </div>
            <div class="input-group">
                <label>Bike Model</label>
                <input type="text" name="bike_model" required placeholder="e.g. Hero Splendor">
            </div>
            <div class="input-group">
                <label>Vehicle Number</label>
                <input type="text" name="vehicle_no" required placeholder="e.g. UP 16 AB 1234">
            </div>
            <div class="input-group">
                <label>Licence / ID Number</label>
                <input type="text" name="doc_id" required placeholder="Enter Driving Licence / Aadhaar">
            </div>
            <button type="submit" class="btn-submit">Submit Registration</button>
        </form>
    </div>
</body>
</html>
""")

@app.route("/api/register-captain", methods=["POST"])
def register_captain():
    name = request.form.get("name")
    phone = request.form.get("phone")
    bike_model = request.form.get("bike_model")
    vehicle_no = request.form.get("vehicle_no")
    doc_id = request.form.get("doc_id")
    
    init_excel()
    wb = load_workbook(CAPTAINS_FILE)
    ws = wb["Captains"]
    ws.append([name, phone, bike_model, vehicle_no, doc_id, "Pending", "₹0", "5.0 ⭐"])
    wb.save(CAPTAINS_FILE)
    return "<script>alert('⏳ Registration Submitted! Admin verification required.');window.location.href='/';</script>"

@app.route("/api/proxy-search")
def proxy_search():
    query = request.args.get("q", "").strip().lower()
    if not query: return jsonify([])
    results = []
    for p in PRESET_LANDMARKS:
        if query in p["name"].lower():
            results.append({"display_name": p["name"], "lat": str(p["lat"]), "lon": str(p["lon"])})
    try:
        headers = {'User-Agent': 'RiderApp/1.0 (sonu@rider.com)'}
        search_q = query if ("noida" in query or "india" in query) else f"{query}, Noida, Uttar Pradesh, India"
        url = f"https://nominatim.openstreetmap.org/search?format=json&q={urllib.parse.quote(search_q)}&countrycodes=in&limit=5"
        resp = requests.get(url, headers=headers, timeout=5)
        api_data = resp.json()
        if isinstance(api_data, list):
            for item in api_data:
                results.append({"display_name": item.get("display_name"), "lat": item.get("lat"), "lon": item.get("lon")})
    except Exception:
        pass
    return jsonify(results[:6])

@app.route("/api/proxy-reverse")
def proxy_reverse():
    lat = request.args.get("lat")
    lon = request.args.get("lon")
    try:
        headers = {'User-Agent': 'RiderApp/1.0 (sonu@rider.com)'}
        url = f"https://nominatim.openstreetmap.org/reverse?lat={lat}&lon={lon}&format=json"
        resp = requests.get(url, headers=headers, timeout=5)
        return jsonify(resp.json())
    except Exception:
        return jsonify({})

@app.route("/api/send-sms-otp", methods=["POST"])
def api_send_sms_otp():
    data = request.json
    phone = data.get("phone")
    otp_code = str(random.randint(1000, 9999))
    USER_OTP_STORE[phone] = otp_code
    return jsonify({"status": "success", "otp": otp_code})

@app.route("/api/verify-sms-otp", methods=["POST"])
def api_verify_sms_otp():
    data = request.json
    phone = data.get("phone")
    otp = data.get("otp")
    if USER_OTP_STORE.get(phone) == otp: return jsonify({"status": "success"})
    return jsonify({"status": "failed"})

@app.route("/api/apply-coupon", methods=["POST"])
def api_apply_coupon():
    data = request.json
    coupon = data.get("coupon")
    if coupon in VALID_COUPONS:
        return jsonify({"status": "success", "discount": VALID_COUPONS[coupon]})
    return jsonify({"status": "invalid"})

@app.route("/api/book", methods=["POST"])
def api_book():
    data = request.json
    name = data.get("name")
    phone = data.get("phone")
    pickup = data.get("pickup")
    drop = data.get("drop")
    service = data.get("service")
    fare = data.get("fare")
    payment_mode = data.get("payment_mode")
    lat = data.get("lat")
    lng = data.get("lng")
    coupon = data.get("coupon", "None")

    ride_otp = str(random.randint(1000, 9999))
    b_index = save_to_excel(name, phone, pickup, drop, service, fare, payment_mode, "Searching...", ride_otp, lat, lng, coupon)

    return jsonify({"status": "success", "booking_index": b_index})

@app.route("/api/booking-status/<int:index>")
def api_booking_status(index):
    init_excel()
    wb = load_workbook(EXCEL_FILE)
    ws = wb["Bookings"]
    row_num = index + 2
    if row_num <= ws.max_row:
        status_val = ws.cell(row=row_num, column=9).value
        captain_assigned = ws.cell(row=row_num, column=8).value
        fare_val = ws.cell(row=row_num, column=6).value
        payment_val = ws.cell(row=row_num, column=7).value
        otp_val = ws.cell(row=row_num, column=10).value

        whatsapp_msg = f"⚡ *RIDER BOOKING ACCEPTED*\n🛵 *Captain:* {captain_assigned}\n🔑 *OTP:* {otp_val}"
        whatsapp_url = f"https://wa.me/{OWNER_WHATSAPP_NUMBER}?text={urllib.parse.quote(whatsapp_msg)}"

        return jsonify({
            "status": status_val,
            "fare": fare_val,
            "payment_mode": payment_val,
            "otp": otp_val,
            "whatsapp_url": whatsapp_url,
            "captain": {"name": captain_assigned, "phone": "9876543210", "bike": "Hero Splendor (UP 16 AB 1234)", "rating": "4.9 ⭐"}
        })
    return jsonify({"status": "Pending"})

@app.route("/api/submit-rating", methods=["POST"])
def api_submit_rating():
    data = request.json
    index = data.get("index")
    rating = data.get("rating")
    
    init_excel()
    wb = load_workbook(EXCEL_FILE)
    ws = wb["Bookings"]
    row_num = index + 2
    if row_num <= ws.max_row:
        ws.cell(row=row_num, column=13, value=rating)
        wb.save(EXCEL_FILE)
        return jsonify({"status": "success"})
    return jsonify({"status": "failed"})

if __name__ == "__main__":
    init_excel()
    app.run(debug=True, port=5000)