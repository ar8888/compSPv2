file_name = "logs.txt"

def clear_log():
    file = open(file_name, "w")
    file.close()

def write_log(message):
    with open(file_name, "a") as file:
        print(message)
        file.write(message+"\n")
