import gspread 
import pandas as pd
import matplotlib.pyplot as plt
from oauth2client.service_account import ServiceAccountCredentials
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib import colors
from reportlab.lib.units import inch
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import SimpleDocTemplate, Paragraph, Table, TableStyle, Image

# Google Sheets API Setup
SCOPE = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
CREDENTIALS_FILE = "credentials.json"  # Your API credentials file

# Connect to Google Sheets
credentials = ServiceAccountCredentials.from_json_keyfile_name(CREDENTIALS_FILE, SCOPE)
client = gspread.authorize(credentials)

# Fetch Google Sheet Data
SHEET_NAME = "Aignosis form (Responses)"  
sheet = client.open(SHEET_NAME).sheet1
data = sheet.get_all_records()

# Convert to DataFrame
df = pd.DataFrame(data)

SOCIAL_PARAMETERS = [
    '"Crows", Laugh', 'Balance Head',
       'Grasps Object within reach', 'Reaches for familiar persons',
       'Rolls over', 'Reaches for nearby objects', 'Occupies self upright',
       'Sits unsupported', 'Pulls self upright', '"talks", imitates sound',
       'Drinks from cup or glass assisted', 'Moves about on floor',
       'Grasps with thumb and finger', 'Demands personal attention',
       'Stands alone', 'Does not drool', 'Follows simple instructions',
       'Walks about room unattended', 'Marks with pencil or crayon',
       'Masticates solid or semi solid food', 'Pulls off clothes',
       'Transfers Object', 'Overcomes simple objects',
       'Fetches or carries similar objects', 'Drinks from cup or glass',
       'Walks without support', 'Plays with other children',
       'Eats with own hands', 'Goes about hours or yard',
       'Discriminated edible substances from non edible',
       'Uses names of familiar objects', 'Walks upstairs unassisted',
       'Unwraps sweets, chocolates', 'Talks in short sentences',
       'Signals to go to toilet', 'Initiates own play activities',
       'Removes shirt to or frock if unbuttoned', 'Eats with spoon/hands',
       'gets drink(water) unassisted', 'Dries own hands',
       'Avoids simple hazards', 'Puts on short or frock unassisted',
       'Can do paper folding', 'Relates experience',
       'Walks downstairs, one step at a time',
       'Plays co-operatively at kindergarten level', 'Buttons shirt or frock',
       'Helps at little household tasks', '"Performs" for others',
       'Washes hand unaided', 'Cares for self at toilet',
       'Washes face unassisted', 'Goes about neighborhood unattended',
       'Dresses self expect for trying',
       'Uses pencil or crayon or chalk for drawing',
       'Plays competitive exercise games', 'Uses hoops or uses knife',
       'Prints(writes) simple words',
       'Plays simple games which require taking turns',
       'Is trusted with money', 'Goes to school unattended',
       'Mixes rice properly unassisted', 'use pencil or chalk for drawing',
       'Bathes self unassisted', 'Goes to bed unassisted',
       'Can differentiate between AM & PM', 'Helps himself during meals',
       'Understands and keeps family secrets',
       'Participants in pre-adolescent', 'Combs or burses hair',
       'Uses tools or utensils', 'Does routine household tasks',
       'Reads on own initiative', 'Bathes self unaided',
       'Cares for self at meals', 'Makes minor purchase',
       'Goes about town freely',
       'Distinguishes between friends any play mates',
       'Makes independent choice of shops', 'Does small remunerative work',
       'Follows local current events', 'Does simple creative work',
       'Is left to care for self or others', 'Enjoys reading books',
       'Plays difficult games', 'Exercises complete care of dress',
       'Buys own clothing accessories',
       'Engages of adolescent group activities',
       'Performs responsible routine chores'
]


# Define Categories and Mapping
Domains = {
            "Self Help General (SHG)": [
                {"item": "Does not drool", "age": 1.0},
                {"item": "Signals to go to toilet", "age": 2.0},
                {"item": "Dries own hands", "age": 2.5},
                {"item": "Avoids simple hazards", "age": 3.0},
                {"item": "Washes hand unaided", "age": 3.5},
                {"item": "Cares for self at toilet", "age": 4.0},
                {"item": "Washes face unassisted", "age": 4.0},
                {"item": "Bathes self unassisted", "age": 6.0},
                {"item": "Goes to bed unassisted", "age": 6.5},
                {"item": "Bathes self unaided", "age": 7.0},
                {"item": "Cares for self at meals", "age": 7.5}
            ],
            "Self Help Eating (SHE)": [
                {"item": "Drinks from cup or glass assisted", "age": 0.9},
                {"item": "Masticates solid or semi solid food", "age": 1.2},
                {"item": "Drinks from cup or glass", "age": 1.5},
                {"item": "Eats with own hands", "age": 1.5},
                {"item": "Discriminated edible substances from non edible", "age": 1.8},
                {"item": "Unwraps sweets, chocolates", "age": 2.0},
                {"item": "Eats with spoon/hands", "age": 2.5},
                {"item": "Gets drink(water) unassisted", "age": 2.5},
                {"item": "Uses hoops or uses knife", "age": 5.0},
                {"item": "Mixes rice properly unassisted", "age": 5.5},
                {"item": "Helps himself during meals", "age": 6.0}
            ],
            "Self Help Dressing (SHD)": [
                {"item": "Pulls off clothes", "age": 1.2},
                {"item": "Removes shirt to or frock if unbuttoned", "age": 2.5},
                {"item": "Puts on short or frock unassisted", "age": 3.0},
                {"item": "Buttons shirt or frock", "age": 3.5},
                {"item": "Dresses self expect for trying", "age": 4.0},
                {"item": "Combs or burses hair", "age": 6.5},
                {"item": "Exercises complete care of dress", "age": 8.5},
                {"item": "Buys own clothing accessories", "age": 9.0}
            ],
            "Self Direction (SD)": [
                {"item": "Demands personal attention", "age": 1.0},
                {"item": "Initiates own play activities", "age": 2.5},
                {"item": "Can differentiate between AM & PM", "age": 6.0},
                {"item": "Understands and keeps family secrets", "age": 6.5},
                {"item": "Makes minor purchase", "age": 7.5},
                {"item": "Makes independent choice of shops", "age": 8.0},
                {"item": "Is left to care for self or others", "age": 8.5},
                {"item": "Buys own clothing accessories", "age": 9.0}
            ],
            "Occupation (OCC)": [
                {"item": "Occupies self upright", "age": 0.6},
                {"item": "Marks with pencil or crayon", "age": 1.2},
                {"item": "Transfers Object", "age": 1.2},
                {"item": "Overcomes simple objects", "age": 1.3},
                {"item": "Fetches or carries similar objects", "age": 1.4},
                {"item": "Can do paper folding", "age": 3.0},
                {"item": "Helps at little household tasks", "age": 3.5},
                {"item": "Uses pencil or crayon or chalk for drawing", "age": 4.0},
                {"item": "Prints(writes) simple words", "age": 5.0},
                {"item": "Is trusted with money", "age": 5.5},
                {"item": "Use pencil or chalk for drawing", "age": 5.5},
                {"item": "Uses tools or utensils", "age": 6.5},
                {"item": "Does routine household tasks", "age": 6.5},
                {"item": "Reads on own initiative", "age": 7.0},
                {"item": "Does small remunerative work", "age": 8.0},
                {"item": "Does simple creative work", "age": 8.5},
                {"item": "Enjoys reading books", "age": 8.5},
                {"item": "Performs responsible routine chores", "age": 10.0}
            ],
            "Communication (COM)": [
                {"item": "Crows", "age": 0.2},
                {"item": "Laugh", "age": 0.3},
                {"item": "\"talks\", imitates sound", "age": 0.8},
                {"item": "Follows simple instructions", "age": 1.0},
                {"item": "Uses names of familiar objects", "age": 1.8},
                {"item": "Talks in short sentences", "age": 2.0},
                {"item": "Relates experience", "age": 3.0},
                {"item": "\"Performs\" for others", "age": 3.5},
                {"item": "Follows local current events", "age": 8.0}
            ],
            "Locomotion (LOC)": [
                {"item": "Balance Head", "age": 0.3},
                {"item": "Rolls over", "age": 0.5},
                {"item": "Sits unsupported", "age": 0.6},
                {"item": "Pulls self upright", "age": 0.7},
                {"item": "Moves about on floor", "age": 0.9},
                {"item": "Stands alone", "age": 1.0},
                {"item": "Walks about room unattended", "age": 1.2},
                {"item": "Walks without support", "age": 1.5},
                {"item": "Goes about hours or yard", "age": 1.8},
                {"item": "Walks upstairs unassisted", "age": 2.0},
                {"item": "Walks downstairs, one step at a time", "age": 3.0},
                {"item": "Goes about neighborhood unattended", "age": 4.0},
                {"item": "Goes to school unattended", "age": 5.5},
                {"item": "Goes about town freely", "age": 7.5}
            ],
            "Socialization (SOC)": [
                {"item": "Grasps Object within reach", "age": 0.4},
                {"item": "Reaches for familiar persons", "age": 0.4},
                {"item": "Reaches for nearby objects", "age": 0.5},
                {"item": "Grasps with thumb and finger", "age": 0.9},
                {"item": "Plays with other children", "age": 1.5},
                {"item": "Plays co-operatively at kindergarten level", "age": 3.0},
                {"item": "Plays competitive exercise games", "age": 5.0},
                {"item": "Plays simple games which require taking turns", "age": 5.0},
                {"item": "Participants in pre-adolescent", "age": 6.5},
                {"item": "Distinguishes between friends any play mates", "age": 7.5},
                {"item": "Plays difficult games", "age": 8.5},
                {"item": "Engages of adolescent group activities", "age": 9.5}
            ]
        }

SA_MAPPING = {
    (0, 5): 2,
    (6, 10): 4,
    (11, 15): 6,
    (16, 20): 8,
    (21, 25): 10,
    (26, 30): 12,
    (31, 40): 14,
    (41, 50): 16
}

def get_social_age(total_score):
    for score_range, sa in SA_MAPPING.items():
        if score_range[0] <= total_score <= score_range[1]:
            return sa
    return 18  # Default max SA if score is very high

def calculate_vsms_scores(row, ca):
    scores = {}
    total_score = 0
    response_count = {"yes": 0, "no": 0, "could've": 0}

    for param in SOCIAL_PARAMETERS:
        response = row.get(param, "").strip().lower()
        if response == "yes":
            score = 1
            response_count["yes"] += 1
        elif response == "no":
            score = 0
            response_count["no"] += 1
        elif response == "could've":
            score = 0.5
            response_count["could've"] += 1
        else:
            score = 0

        total_score += score
        scores[param] = score

    sa = get_social_age(total_score)
    sq = (sa / ca) * 10 if ca > 0 else 0

    if sq >= 110:
        maturity = "Above Average"
    elif sq >= 90:
        maturity = "Average"
    elif sq >= 70:
        maturity = "Borderline"
    elif sq >= 50:
        maturity = "Mild"
    elif sq >= 35:
        maturity = "Moderate"
    elif sq >= 20:
        maturity = "Severe"
    else:
        maturity = "Profound"

    return total_score, sa, sq, maturity, scores, response_count


# Compute Scores Based on Categories
def compute_scores(row):
    category_scores = {cat: 0 for cat in Domains.keys()}
    response_count = {"yes": 0, "no": 0, "could've": 0}
    
    for category, params in Domains.items():
        for param in params:
            response = row.get(param, "").strip().lower()
            score = 1 if response == "yes" else 0.5 if response == "could've" else 0
            category_scores[category] += score
            
            if response in response_count:
                response_count[response] += 1
    
    return category_scores, response_count

#Generate Bar Chart
def generate_bar_chart(category_scores, name):
    plt.figure(figsize=(8, 4))
    plt.bar(category_scores.keys(), category_scores.values(), color="skyblue")
    plt.xlabel("Categories")
    plt.ylabel("Score")
    plt.title(f"Category-wise Social Scores for {name}")
    plt.xticks(rotation=45, ha="right")
    plt.tight_layout()
    
    file_path = f"{name}_bar_chart.png"
    plt.savefig(file_path)
    plt.close()
    return file_path
    

# Generate Pie Chart
def generate_pie_chart(response_count, name):
    labels = response_count.keys()
    sizes = list(response_count.values())
    
    if sum(sizes) == 0:
        print(f"Skipping pie chart for {name} due to no valid responses.")
        return None  # No data
    
    colors = ["green", "red", "orange"]
    plt.figure(figsize=(5, 5))
    plt.pie(sizes, labels=labels, autopct="%1.1f%%", colors=colors, startangle=140)
    plt.title(f"Response Distribution for {name}")
    
    file_path = f"{name}_pie_chart.png"
    plt.savefig(file_path)
    plt.close()
    return file_path

# Generate PDF Report
def generate_pdf(report_data, user_name):
    filename = f"{user_name}_VSMS_Report.pdf"
    doc = SimpleDocTemplate(filename, pagesize=A4)
    elements = []
    styles = getSampleStyleSheet()
    
    elements.append(Paragraph("<b>VSMS Category-wise Assessment Report</b>", styles["Title"]))
    
    for entry in report_data:
        elements.append(Paragraph(f"<b>Name:</b> {entry['Name']}", styles["Normal"]))
        elements.append(Paragraph(f"<b>Email:</b> {entry['Email']}", styles["Normal"]))
        elements.append(Paragraph(f"<b>Chronological Age (CA):</b> {entry['Chronological Age']}", styles["Normal"]))
       # elements.append(Paragraph(f"<b>Total Score:</b> {entry['Total Score']}", styles["Normal"]))
        elements.append(Paragraph(f"<b>Social Age (SA):</b> {entry['Social Age']}", styles["Normal"]))
        elements.append(Paragraph(f"<b>Social Quotient (SQ):</b> {entry['Social Quotient']:.2f}", styles["Normal"]))
       # elements.append(Paragraph(f"<b>Maturity Level:</b> {entry['Maturity Level']}", styles["Normal"]))

        # Table Data
        table_data = [["Category", "Score"]] + [[cat, score] for cat, score in entry["Category Scores"].items()]
        table = Table(table_data, colWidths=[300, 100])
        table.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.grey),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.whitesmoke),
            ("ALIGN", (0, 0), (-1, -1), "CENTER"),
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
            ("GRID", (0, 0), (-1, -1), 1, colors.black),
        ]))
        elements.append(table)
        elements.append(Paragraph("<br/>", styles["Normal"]))
        
        elements.append(Image(entry["Bar Chart"], width=400, height=200))
        if entry["Pie Chart"]:
            elements.append(Image(entry["Pie Chart"], width=300, height=300))
        
        elements.append(Paragraph("<br/><br/>", styles["Normal"]))
    
    doc.build(elements)
    print(f"PDF Report Generated: {filename}")

# Main execution flow
def main():
    user_name = input("Enter the name of the user to generate the report: ")
    
    # Filter data for the specific user
    filtered_df = df[df["Name"].str.lower() == user_name.lower()]
    
    if filtered_df.empty:
        print(f"No records found for {user_name}.")
        return
    
    # Process only the filtered user's data
    report_data = []
    for _, row in filtered_df.iterrows():
        ca = int(row.get("Actual Age in Numbers", 0))
        total_score, sa, sq, maturity, scores, response_count = calculate_vsms_scores(row, ca)
        category_scores, response_count = compute_scores(row)
        bar_chart_path = "bar_chart.png"
        pie_chart_path = generate_pie_chart(response_count, row.get("Name", "Unknown"))
        
        report_data.append({
            "Name": row.get("Name", "Unknown"),
            "Email": row.get("Email Address", "Unknown"),
            "Chronological Age": ca,
            "Total Score": total_score,
            "Social Age": sa,
            "Social Quotient": sq,
            "Maturity Level": maturity,
            "Scores": scores,
            "Category Scores": category_scores,
            "Bar Chart": bar_chart_path,
            "Pie Chart": pie_chart_path
        })
    
    # Generate report only if data was found
    if report_data:
        generate_pdf(report_data, user_name)
    else:
        print(f"No valid data found for {user_name}.")

if __name__ == "__main__":
    main()