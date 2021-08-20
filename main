""" A simple application for Customer Helpdesk Team, utilizing
CRUD, ETL and OOP

Platform and drivers in use:
SQL Server, Python, PYODBC, Jupyter Notebook

A databse with multiple tables were created in SQL Server. Python 
was used with the help of PYODBC driver to retrieve the data from
the tables. This application can be used to add, view, query, and 
edit data. Delete data wasn't initiated since all the customer issues
were to save, rather change the status to 'closed' instead.
A graphical user interface (GUI) was created, and a log-in window
validating user credentials

A project from Udemy course 'Database Programming with Python' by 
Zak Ruvalcaba (Learn how to integrate free and enterprise databases 
into your Python workflow including SQLite, MySQL, and SQL Server)

Name: Ifti Mustafa.
"""


import pyodbc
from contextlib import closing
import datetime

conn = None

def connect():
    global conn
    if not conn:
        conn = pyodbc.connect('Driver={ODBC Driver 17 for SQL Server};'
                      'Server=DESKTOP-1A5AN2K\SQLEXPRESS;'
                      'Database=newline_toronto;'
                      'Trusted_Connection=yes;')

def close():
    if conn:
        conn.close()

def make_ticket(row):
    """ Creating a function to call the Ticket class, to create object """
    
    return Ticket(row.ticketid, row.customername, row.customeremail, row.submitteddate,
                 row.employeeid, row.solutionid, row.statusid, row.issue)

def get_open_tickets():
    """ To query and show all the open tickets with multiple fields 
    from different tables:
    we have JOINed between employees and tickets, 
                   between tickets and solutions,
                   between tickets and status """
    
    query = ("SELECT tickets.ticketid, status.status, solutions.solution, " +
            "employees.name, tickets.customername, tickets.customeremail, " +
            "tickets.submitteddate, tickets.issue, tickets.employeeid, " +
            "tickets.solutionid, tickets.statusid " +
            "FROM employees INNER JOIN " +
            "tickets ON employees.employeeid = tickets.employeeid INNER JOIN " +
            "solutions ON tickets.solutionid = solutions.solutionid INNER JOIN " +
            "status on tickets.statusid = status.statusid " +
            "WHERE tickets.statusid < 3")
    with closing(conn.cursor()) as cursor:
        cursor.execute(query)
        results = cursor.fetchall()
    tickets = []
    for row in results:
        tickets.append(row)
    return tickets

def get_ticket_detail(ticketid):
    query = ("SELECT * FROM tickets WHERE ticketid = ?")
    with closing(conn.cursor()) as cursor:
        cursor.execute(query,(ticketid,))
        result = cursor.fetchone()
    ticket = make_ticket(result)
    return ticket

def add_ticket(ticket):
    sql = ("INSERT INTO tickets (ticketid, customername, customeremail, submitteddate, employeeid, solutionid, statusid, issue) VALUES (?,?,?,?,?,?,?,?)")
    with closing(conn.cursor()) as cursor:
        cursor.execute(sql,(ticket.ticketid, ticket.customername, ticket.customeremail, ticket.submitteddate, ticket.employeeid, ticket.solutionid, ticket.statusid, ticket.issue))
        conn.commit()

def update_ticket(statusid, ticketid):
    sql = ("UPDATE tickets SET statusid=? WHERE ticketid = ?")
    with closing(conn.cursor()) as cursor:
        cursor.execute(sql,(statusid, ticketid))
        conn.commit()
    
def employee_info():
    query =  "SELECT * FROM employee_roles"
    with closing(conn.cursor()) as cursor:
        cursor.execute(query)
        results = cursor.fetchall()
    return results

def login(username, password):
    """ Checking credintials for log in. If the user and password matches
    then the 'COUNT' method will give us total of 1, therefore satisfying
    the condition and returning 'True' """
    
    query = "SELECT COUNT(*) FROM employees WHERE username = ? AND password = ?"
    with closing(conn.cursor()) as cursor:
        cursor.execute(query,(username, password))
        result = cursor.fetchone()
    if result[0] > 0:
        return True
    else:
        return False

class Ticket:
    """ class method to create each ticket an object """
    
    def __init__(self, ticketid=0, customername = None, customeremail = None,
                submitteddate = None, employeeid = 0, solutionid = 0,
                statusid = 0, issue = None):
        self.ticketid = ticketid
        self.customername = customername
        self.customeremail = customeremail
        self.submitteddate = submitteddate
        self.employeeid = employeeid
        self.solutionid = solutionid
        self.statusid = statusid
        self.issue = issue

def display_menu():
    """ Graphical User Interface, main display """
    
    print("NEWLINE TORONTO LTD. (Admin)",'\n')
    print("COMMAND MENU")
    print("-" * 75)
    print("View - View all open tickets")
    print("Det - Detail view by ticket number")
    print("Add - Add new ticket")
    print("Upd - Update existing ticket")
    print("Exit - Exit the system")
    print("-" * 75)
    
def view_tickets():
    print("NEWLINE TORONTO CUSTOMER OPEN TICKETS LIST")
    print("-" * 75)
    line_format = "{:15s} {:15s} {:15s} {:15s} {:15s}"
    print(line_format.format("Ticket ID", "Full Name", "Date", "Solution", "Status"))
    print("-" * 75)
    new = get_open_tickets()
    for i in new:
        print(line_format.format(str(i.ticketid), i.customername, str(i.submitteddate), i.solution, i.status))
    print("-" * 75)
    
def view_ticket_detail():
    ticket_id = int(input("Enter the ticket id: "))
    i = get_ticket_detail(ticket_id)
    print("NEWLINE TORONTO CUSTOMER TICKETS LIST")
    print("-" * 75)
    line_format = "{:15s} {:15s} {:15s} {:15s} {:15s}"
    line_format2 = "{:15s} {:15s} {:50s}"
    print(line_format.format("Ticket ID", "Customer Name", "Email", "Date", "Employee ID"))
    print("-" * 75)
    print(line_format.format(str(i.ticketid), i.customername, i.customeremail, str(i.submitteddate), str(i.employeeid)))
    print("-" * 75,'\n')
    print(line_format2.format("Solution ID", "Status ID", "Issue"))
    print("-" * 75)
    print(line_format2.format(str(i.solutionid), str(i.statusid), i.issue))
    print("-" * 75)
    
def add_ticket_ui():
    ticketid = int(input("Enter ticket ID: "))
    customername = input("Enter customer full name: ")
    customeremail = input("Enter email address: ")
    submitteddate = input("Enter date (YYYY-MM-DD): ")
    for i in employee_info():
        print(i[0], i.name, i.role)
    employeeid = int(input("Choose employee id from list above: "))
    solutionid = int(input("Enter solution id (1:vProspect, 2:vConvert, 3:vRetain): "))
    statusid = int(input("Enter status id (1:Open, 2: In Progress, 3:Closed): "))
    issue = input("Enter issue details: ")
    new = Ticket(ticketid=ticketid, customername = customername, customeremail = customeremail,
                submitteddate = submitteddate, employeeid = employeeid, solutionid = solutionid,
                statusid = statusid, issue = issue)
    add_ticket(new)
    print("New record for ", customername, " has been created successfully.")
    
def update_ticket_ui():
    ticketid = int(input("Enter ticket ID: "))
    statusid = int(input("Enter new status ID: "))
    i = get_ticket_detail(ticketid)
    print(f"Ticket ID you've choosen is {i.ticketid} with status id {i.statusid} and customer name is {i.customername}")
    reply = str(input("Are you sure you want to change the status of the above customer? (y/n)"))
    if reply.lower().strip() == 'y':
        update_ticket(statusid, ticketid)
        print("customer record has been updated successfully")
    else:
        print("Customer status wasn't updated")

def main():
    """ Initiating database connection. Checking log in credentials.
    Calling display_menu (GUI), offering multiple user choices
    to view, detail view, add, update, and exit the application.
    Finally closing the connection """
    
    connect()
    while True:
        print("NEWLINE TORONTO HELPDESK LOG IN")
        print("-" * 75)
        username = input("Enter Username: ")
        password = input("Enter Password: ")
        if login(username,password):
            break
        else:
            print("Invalid credintials, please try again.")
    
    display_menu()
    
    while True:
        command = input("Enter choice: ").lower()
        if command == "view":
            view_tickets()
        elif command == "det":
            view_ticket_detail()
        elif command == "add":
            add_ticket_ui()
        elif command == "upd":
            update_ticket_ui()
        elif command == "exit":
            break
        else:
            print("Wrong input, try again!")
            display_menu()
    close()

if __name__ == "__main__":
    main()
