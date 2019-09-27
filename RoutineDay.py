class DailyRoutine:
    def __init__(self, dayNumber, _class):
        self.dayNumber = dayNumber
        self._class = _class

        # 0 - saturday
        # 1 - Sunday
        # 2 - Monday
        # 3 - Tuesday
        # 4 - Wednesday
        # 5 - Thursday


class SingleClass:
    def __init__(self, room_number, course_code, assigned_teacher):
        self.room_number = room_number
        self.course_code = course_code
        self.assigned_teacher = assigned_teacher

    def get_info(self):
        print("SingleClass: ", self.room_number, self.course_code, self.assigned_teacher)
    # this method is only being used for testing purpose
