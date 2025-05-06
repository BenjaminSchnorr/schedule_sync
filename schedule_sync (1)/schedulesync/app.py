import nest_asyncio
nest_asyncio.apply()
from flask import Flask, render_template, redirect, jsonify, session, request
import json
import secrets
import pandas as pd
import requests
import asyncio

import async_timeout
import threading
from msgraph.generated.users.users_request_builder import UsersRequestBuilder
from azure.identity import AuthorizationCodeCredential

from msgraph import GraphServiceClient
from msgraph.generated.models.attendee import Attendee
from msgraph.generated.models.attendee_type import AttendeeType
from msgraph.generated.models.day_of_week import DayOfWeek
from msgraph.generated.models.email_address import EmailAddress
from msgraph.generated.models.patterned_recurrence import PatternedRecurrence
from msgraph.generated.models.recurrence_pattern import RecurrencePattern
from msgraph.generated.models.recurrence_pattern_type import RecurrencePatternType
from msgraph.generated.models.recurrence_range import RecurrenceRange
from msgraph.generated.models.recurrence_range_type import RecurrenceRangeType
from msgraph.generated.users.item.events.events_request_builder import EventsRequestBuilder
from kiota_abstractions.base_request_configuration import RequestConfiguration
from msgraph.generated.models.event import Event
from msgraph.generated.models.item_body import ItemBody
from msgraph.generated.models.body_type import BodyType
from msgraph.generated.models.date_time_time_zone import DateTimeTimeZone
from msgraph.generated.models.location import Location

app = Flask(__name__)
app.secret_key = secrets.token_hex(16)
progress = 0
response = ""
error_message = ""
undergrad_options_url = 'https://apps.western.edu/cs/undergrad_options'
graduate_options_url = 'https://apps.western.edu/cs/grad_options'
lock = threading.Lock()
auth_code = ""

client = None

with open("passwords.json", "r") as file:
    passwords = json.load(file)
    admin_password = passwords["admin_password"]
    client_id = passwords["client_id"]
    tenant_id = passwords["tenant_id"]
    client_secret = passwords["client_secret"]


@app.route("/change_password", methods=["POST", "GET"])
def change_password():
    if "logged_in" not in session or session["logged_in"] == False:
        return redirect("/login")
    if request.method == "POST":
        global admin_password
        current_password = request.form.get("current_password")
        new_password = request.form.get("new_password")
        confirm_password = request.form.get("confirm_password")
        if current_password != admin_password:
            message = "Current Password Incorrect"
        elif new_password != confirm_password:
            message = "New Password and Confirm Password do not match"
        else:
            message = "Password Changed, logout and log back in with your new password"
            with open("passwords.json", "r") as file:
                passwords = json.load(file)
            with open("passwords.json", "w") as file:
                passwords["admin_password"] = new_password
                json.dump(passwords, file)


            admin_password = new_password

        return render_template("/change_password.html", message=message)
    if request.method == "GET":
        return render_template("change_password.html", message="")


@app.route("/microsoft_authentication", methods=["POST", "GET"])
def microsoft_authentication():
    if request.method == "POST":
        global auth_code, client_id, tenant_id, client_secret

        auth_code = request.form.get("code")
        credentials = AuthorizationCodeCredential(
        tenant_id=tenant_id,
        client_id=client_id,
        authorization_code=auth_code,
        redirect_uri="https://benschnorr.pythonanywhere.com/microsoft_authentication",
        client_secret=client_secret
    )
    global client
    scopes = ["Calendars.ReadWrite", "User.ReadBasic.All"]
    client = GraphServiceClient(credentials=credentials, scopes=scopes)


    return redirect("/")

@app.route("/")
def home():
    global error_message
    response = requests.get(undergrad_options_url).json()
    undergrad_terms = response["Standard_Term"]
    response = requests.get(graduate_options_url).json()
    grad_terms = response["Standard_Term"]
    undergrad_term_length = len(undergrad_terms)
    grad_term_length = len(grad_terms)
    if undergrad_term_length < grad_term_length:
        terms = undergrad_terms
    else:
        terms = grad_terms

    if "logged_in" not in session or session["logged_in"] == False:
        return redirect("/login")
    else:
        return render_template('index.html', terms=terms)


@app.route("/logout")
def logout():
    session["logged_in"] = False
    return redirect("/")


@app.route("/login", methods=["POST", "GET"])
def login():
    if request.method == "POST":
        password = request.form.get("password")
        if password == admin_password:
            session["logged_in"] = True
            scopes = ["https://graph.microsoft.com/Calendars.ReadWrite","https://graph.microsoft.com/User.ReadBasic.All"]
            authentication_request = "https://login.microsoftonline.com/0ef6279f-a082-4fe7-b7a4-4791c9ed6d1c/oauth2/v2.0/authorize?client_id=08aa281e-db34-4358-8617-95b09cdb335f&response_type=code&redirect_uri=https://benschnorr.pythonanywhere.com/microsoft_authentication&response_mode=form_post&scope=" + str(scopes[0]) +"%20" + str(scopes[1])
            request_for_code = requests.get(authentication_request)
            return redirect(authentication_request)
        else:
            session["logged_in"] = False
            return render_template('login.html')
    else:
        if "logged_in" in session and session["logged_in"]:
            return redirect("/")
        return render_template("login.html")



@app.route("/check_progress")
def check_progress():
    if "logged_in" not in session or session["logged_in"] == False:
        return redirect("/login")
    with lock:
        global progress
        global error_message

    return jsonify({"progress": progress, "error_message": str(error_message)})

@app.route("/start_retrieval", methods=['POST'])
def start_retrieval():
    if "logged_in" not in session or session["logged_in"] == False:
        return redirect("/login")
    term = request.form["term"]
    with lock:
        global progress
        global error_message
        progress = 0
        error_message = ""
    thread = threading.Thread(target=retrieve_schedules, args=(term,))
    thread.start()
    return jsonify({'message': 'retrieval process started'})


@app.route("/start_meeting_invites")
def start_meeting_invites():
    if "logged_in" not in session or session["logged_in"] == False:
        return redirect("/login")
    with lock:
        global progress
        global error_message
        progress = 0
        error_message = ""
    thread = threading.Thread(target=generate_meeting_invites)
    thread.start()

    return jsonify({'message': 'invite generation started'})


def generate_meeting_invites():
    with lock:
        global client


    days_of_week_map = {
        "Mon": DayOfWeek.Monday,
        "Tues": DayOfWeek.Tuesday,
        "Wed": DayOfWeek.Wednesday,
        "Thurs": DayOfWeek.Thursday,
        "Fri": DayOfWeek.Friday
    }

    coroutines = []

    request_configuration = RequestConfiguration()
    request_configuration.headers.add("Prefer", "outlook.timezone=\"America/Denver\"")

    def parse_hours(hours):
        days_of_week, time = hours.split(" | ")
        start_time, end_time = time.split(" - ")
        if "PM" in start_time and "12:" not in start_time:
            if start_time[1] == ":":
                hour = int(start_time[:1])
                hour_military = hour + 12
                hour_str = str(hour_military)
                start_time = start_time.replace(str(hour), hour_str, 1)
            else:
                hour = int(start_time[:2])
                hour_military = hour + 12
                hour_str = str(hour_military)
                start_time = start_time.replace(str(hour), hour_str, 1)
        if "PM" in end_time and "12:" not in end_time:
            if end_time[1] == ":":
                hour = int(end_time[:1])
                hour_military = hour + 12
                hour_str = str(hour_military)
                end_time = end_time.replace(str(hour), hour_str, 1)
            else:
                hour = int(end_time[:2])
                hour_military = hour + 12
                hour_str = str(hour_military)
                end_time = end_time.replace(str(hour), hour_str, 1)

        days_of_week = days_of_week.split("/")
        for i in range(len(days_of_week)):
            day = days_of_week[i]
            if day in days_of_week_map.keys():
                days_of_week[i] = days_of_week_map[day]

        start_time, _ = start_time.split(" ")
        end_time, _ = end_time.split(" ")
        if start_time[1] == ":":
            start_time = "0" + start_time
        if end_time[1] == ":":
            end_time = "0" + end_time
        return start_time, end_time, days_of_week

    async def request(request_body):
        result = await asyncio.wait_for(
            client.me.events.post(request_body, request_configuration=request_configuration), timeout=300)

    def generate_invites(row):
        start_date = row.Start_Date
        end_date = row.End_Date
        location = row.Locations
        course_name = row.Name
        hours = row.Hours
        start_time, end_time, days_of_week = parse_hours(hours)
        email = row["emails"]
        request_body = Event(
            subject=course_name,
            #get rid of this item_body code for production
            body=ItemBody(
                content_type=BodyType.Html,
                content=email,
            ),
            start=DateTimeTimeZone(
                date_time=start_date + "T" + start_time + ":00",
                time_zone="America/Denver",
            ),
            end=DateTimeTimeZone(
                date_time=start_date + "T" + end_time + ":00",
                time_zone="America/Denver",
            ),
            recurrence=PatternedRecurrence(
                pattern=RecurrencePattern(
                    type=RecurrencePatternType.Weekly,
                    interval=1,
                    days_of_week=days_of_week,
                ),
                range=RecurrenceRange(
                    type=RecurrenceRangeType.EndDate,
                    start_date=start_date,
                    end_date=end_date,
                ),
            ),
            location=Location(
                display_name=location,
            ),
            #uncomment this code to send emails to all professors
            # attendees = [
            # 	Attendee(
            # 		email_address = EmailAddress(
            #
            # 			address = email,
            # 		),
            # 		type = AttendeeType.Required,
            # 	),
            # ],
            allow_new_time_proposals=True,
        )
        coroutine = request(request_body)
        coroutines.append(coroutine)

    df = pd.read_csv('schedule_information.csv')
    df = df.dropna(subset=["Hours"])
    df = df.dropna(subset=["Instructor"])
    new_rows = []

    def check_for_labs(row):
        hours = row.Hours
        if "; " in hours:
            lab_hours, class_hours = hours.split("; ")
            row.Hours = class_hours
            new_row = row
            new_row.Hours = lab_hours
            new_rows.append(new_row.to_dict())
        return row

    new_df = df.apply(check_for_labs, axis=1)

    for row in new_rows:
        temp_df = pd.DataFrame([row])
        new_df = pd.concat([new_df, temp_df])

    new_df = new_df.reset_index(drop=False)

    instructors = new_df["Instructor"].unique()

    for instructor in instructors:
        index = new_df["Instructor"] == instructor
        instructor_classes = new_df[index]
        instructor_classes.apply(generate_invites, axis=1)


    async def run_coroutines():
        for routine in coroutines:
            result = await routine
            with lock:
                global progress
                progress += 100 / len(coroutines)

    df.to_csv("schedule_information.csv", index=False)

    try:
        asyncio.run(run_coroutines())
    except Exception as e:
        global error_message
        error_message = "generate_error"


def retrieve_schedules(term):
    undergrad_search_url = 'https://apps.western.edu/cs/undergrad_search'
    undergrad_options_url = 'https://apps.western.edu/cs/undergrad_options'
    graduate_search_url = 'https://apps.western.edu/cs/grad_search'
    graduate_options_url = 'https://apps.western.edu/cs/grad_options'

    columns = ["Hours", "Instructor", "Instructional_Format", "End_Date", "Delivery_Mode", "Name", "Section_Details",
               "Days_of_the_Week", "Start_Date", "Course", "Locations", "Title", "Not_Online"]

    request = {
        "query": {
            "filters": [
                {
                    "field": "Standard_Term",
                    "value": term,
                    "include": "True"
                },
                {
                    "field": "Instructors",
                    "value": "",
                    "include": "True"
                }
            ],
            "searches": [
                {
                    "field": "Course",
                    "value": ""
                },
                {
                    "field": "Prerequisite",
                    "value": ""
                },
                {
                    "field": "Course_Tags",
                    "value": ""
                },
                {
                    "field": "Subject",
                    "value": ""
                },
                {
                    "field": "Instructors",
                    "value": ""
                },
                {
                    "field": "Name",
                    "value": ""
                }
            ]
        }
    }

    response = requests.get(undergrad_options_url).json()
    instructors_temp = response["Instructors"]
    instructors = instructors_temp
    instructor_information = {}

    response = requests.get(graduate_options_url).json()
    instructors_graduate_temp = response["Instructors"]
    instructors_graduate = instructors_graduate_temp
    requests_course_search = len(instructors) + len(instructors_graduate)
    total_requests = requests_course_search * 3.5

    for instructor in instructors:
        global error_message
        try:
            request["query"]["filters"][1]["value"] = instructor
            response = requests.post(undergrad_search_url, json=request, timeout=5)
            instructor_information[instructor] = response.json()
            with lock:
                global progress
                progress += 100.0 / total_requests
        except:
            with lock:
                error_message = "retrieval_error"

    for instructor in instructors_graduate:
        try:
            request["query"]["filters"][1]["value"] = instructor
            response = requests.post(graduate_search_url, json=request, timeout=5)
            instructor_information[instructor + "_grad"] = response.json()
            with lock:
                progress += 100.0 / total_requests
        except:
            with lock:
                error_message = "retrieval_error"



    rows = []

    for instructor, information in instructor_information.items():
        info = json.dumps(information, indent=2)

        for course in information:
            row = []

            for column in columns:
                if column in course.keys():
                    row.append(course[column])
                else:
                    row.append("NA")
            rows.append(row)

    def check_for_labs(row):
        hours = row.Hours
        if "; " in hours:
            lab_hours, class_hours = hours.split("; ")
            row.Hours = class_hours
            new_row = row
            new_row.Hours = lab_hours
            pd.concat([df, new_row])

    df = pd.DataFrame(rows, columns=columns)
    df = df.dropna(subset=["Hours"])
    df = df.dropna(subset=["Instructor"])
    instructors = df["Instructor"].tolist()
    instructors_emails = []
    with lock:
        global client
    async def find_user(client: GraphServiceClient):
        for instructor in instructors:
            if instructor:
                user_name = str(instructor)
                if '\'' in user_name:
                    user_name = user_name.replace('\'', "")
                name_split = user_name.split(" ")
                if len(name_split) == 2:
                    first_name = name_split[0]
                    last_name = name_split[1]
                elif len(name_split) == 3:
                    first_name = name_split[0]
                    last_name = name_split[1] + " " + name_split[2]

                query_params = UsersRequestBuilder.UsersRequestBuilderGetQueryParameters(
                    filter=f"displayName eq '{user_name}'"
                )
                request_configuration = (
                    UsersRequestBuilder.UsersRequestBuilderGetRequestConfiguration(
                        query_parameters=query_params,
                    )
                )
                request_configuration.headers.add("ConsistencyLevel", "eventual")
                response = await client.users.get(request_configuration=request_configuration)

                if not response.value:
                    query_params = UsersRequestBuilder.UsersRequestBuilderGetQueryParameters(

                        filter=f"surName eq '{last_name}' and givenName eq '{first_name}'"

                    )
                    request_configuration = (
                        UsersRequestBuilder.UsersRequestBuilderGetRequestConfiguration(
                            query_parameters=query_params,
                        )
                    )
                    request_configuration.headers.add("ConsistencyLevel", "eventual")
                    response = await client.users.get(request_configuration=request_configuration)
                if not response.value:
                    query_params = UsersRequestBuilder.UsersRequestBuilderGetQueryParameters(
                        filter=f"surName eq '{last_name}'"

                    )
                    request_configuration = (
                        UsersRequestBuilder.UsersRequestBuilderGetRequestConfiguration(
                            query_parameters=query_params,
                        )
                    )
                    request_configuration.headers.add("ConsistencyLevel", "eventual")
                    response = await client.users.get(request_configuration=request_configuration)
                if response.value:
                    with lock:
                        global progress
                        progress += 100.0 / total_requests
                    user = response.value[0]
                    instructors_emails.append(user.mail)
                else:
                    with lock:
                        progress += 100 / total_requests
                    instructors_emails.append("")


    asyncio.run(find_user(client))
    df["emails"] = instructors_emails
    df.to_csv("schedule_information.csv", index=False)
    with lock:
        progress = 100.0

if __name__ == "__main__":
    app.run(debug=True)