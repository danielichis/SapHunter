# get the current date
import datetime
def get_current_date():
    current_date = datetime.datetime.now().date().strftime("%d-%m-%Y")
    return current_date
def learning_breaks(q):
    if q==2:
        print("Break time")
        return
    print("Continue learning")
learning_breaks(3)