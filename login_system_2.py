stored_username = "admin"
stored_password = "Pass@123"

username = input("Please enter the username: ")
password = input("Please enter the password: ")


# Approach #1 (Boolean operator)
if username == stored_username and password == stored_password:
    print("Welcome back!")
elif username != stored_username:
    print("This username is unregistered.")
else:
    print("Password is incorrect.")


# Approach #2 (Nested if-else)
if username == stored_username:
    if password == stored_password:
        print("Welcome back!")
    else:
        print("Password is incorrect.")
else:
    print("This username is unregistered.")


# Approach #3 (Negative approach)
if username != stored_username:
    print("This username is unregistered.")
elif password != stored_password:
    print("Password is incorrect.")
else:
    print("Welcome back!")