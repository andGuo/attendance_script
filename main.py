from typing import Dict, List, NamedTuple

import os
import errno
from openpyxl import load_workbook

# These probably don't need to be changed
TUTORIAL_LIST_FILENAME = "tutorials_merged_20230915"
ATTENDANCE_NAMES_FILE = "bot_input"
APPEND_MODE = True  # if false, will first reset every score to 0

# These needs to be changed
TUTORIAL_NUMBER = 2
MAX_SCORE = 2  # maximum score (int) a student can get for the tutorial


class StudentAttendance(NamedTuple):
    username: str
    sid: int
    score: int


class TakeAttendance:
    def __init__(
        self,
        output_path: str,
        input_path: str,
        tutorial_number: int,
        append_mode: bool = True,
    ) -> None:
        self.output_path: str = output_path
        self.input_path: str = input_path
        self.tutorial_number: str = tutorial_number
        self.append_mode: bool = append_mode

    def _load_student_attendance(self) -> List[str]:
        """
        Load the bot's student attendance from the file specified by self.input_path

        Returns:
            List[str]: a list of strings, each string is a student's username
        """
        try:
            with open(self.input_path, "r") as file:
                lines = file.readlines()
            return lines
        except Exception as e:
            print("ERROR >>> Unable to parse student attendance file!")
            raise Exception(e)

    def _load_tutorial_list(self) -> Dict[str, StudentAttendance]:
        """
        Load the tutorial student list from the file specified by self.output_path

        Returns:
            Dict[str, StudentAttendance]: a dictionary, the key is the username, the value is a StudentAttendance NamedTuple
        """
        try:
            workbook = load_workbook(
                filename=self.output_path, data_only=True, read_only=True
            )
            worksheet = workbook[self.tutorial_number]

            # Determine which columns are sid, username, and score
            for i in range(3):
                col_str = worksheet.cell(0, i).value

                if col_str == "OrgDefinedId":
                    sid_col = i
                elif col_str == "Username":
                    username_col = i
                elif col_str.startswith("Tutorial"):
                    score_col = i
                else:
                    raise Exception("Undefined column name")

            tut_list_dict = {}
            num_rows = 0

            # Count number of non-empty rows
            for row in worksheet:
                if not all(col.value is None for col in row):
                    num_rows += 1

            # Create dict
            for row in range(1, num_rows + 1):
                key = worksheet.cell(row, username_col).value

                if key in tut_list_dict:
                    raise Exception("Duplicate username found")

                if (
                    self.append_mode
                    and worksheet.cell(row, score_col).value is not None
                ):
                    value = StudentAttendance(
                        username=key,
                        sid=worksheet.cell(row, sid_col).value,
                        score=worksheet.cell(row, score_col).value,
                    )
                else:
                    value = StudentAttendance(
                        username=key,
                        sid=worksheet.cell(row, sid_col).value,
                        score=0,
                    )

                tut_list_dict[key] = value

            workbook.close()

            return tut_list_dict

        except Exception as e:
            print("ERROR >>> Unable to parse tutorial file!")
            raise Exception(e)

    def _update_attendance(
        self, attendance: List[str], tutorial_dict: Dict[str, StudentAttendance]
    ) -> List[StudentAttendance]:
        """
        Update the attendance of the students in tutorial_dict with the students in attendance

        Args:
            attendance (List[str]): list of students from the bot's attendance
            tutorial_dict (Dict[str, StudentAttendance]): dictionary of students from the tutorial list

        Returns:
            List[StudentAttendance]: list of students from the tutorial list, updated with the students in attendance
        """
        for username in attendance:
            if username in tutorial_dict:
                old_score = tutorial_dict[username].score
                new_score = 1 if old_score <= 0 else min(old_score, MAX_SCORE)

                tutorial_dict[username] = StudentAttendance(
                    username=username, sid=tutorial_dict[username].sid, score=new_score
                )
                print(f"Updated {username} to {new_score}!")
            else:
                print(f"WARNING >>> {username} not found in tutorial list!")

        return list(tutorial_dict.values())

    def _write_tutorial_list(self, students: List[StudentAttendance]) -> int:
        """
        Write the tutorial list to the file specified by self.output_path

        Args:
            students (List): list of StudentAttendance NamedTuples

        Returns:
            int: number of students written to file
        """
        students.sort(key=lambda x: x.sid)

        try:
            workbook = load_workbook(filename=self.output_path)
            worksheet = workbook[self.tutorial_number]

            # Determine which columns are sid, username, and score
            for i in range(3):
                col_str = worksheet.cell(0, i).value

                if col_str == "OrgDefinedId":
                    sid_col = i
                elif col_str == "Username":
                    username_col = i
                elif col_str.startswith("Tutorial"):
                    score_col = i
                else:
                    raise Exception("Undefined column name")

            # Write to file
            for row, student in enumerate(students, start=1):
                worksheet.cell(row, sid_col).value = student.sid
                worksheet.cell(row, username_col).value = student.username
                worksheet.cell(row, score_col).value = student.score

            workbook.save(filename=self.output_path)
            workbook.close()

            return row - 1

        except Exception as e:
            print("ERROR >>> Unable to write to tutorial file!")
            raise Exception(e)

    def run(self):
        student_attendance = self._load_student_attendance()
        print(f"Loaded {len(student_attendance)} students from {self.input_path}")
        attendance_dict = self._load_tutorial_list()
        print(f"Found {len(attendance_dict)} students from {self.output_path}")
        new_attendance = self._update_attendance(student_attendance, attendance_dict)
        num = self._write_tutorial_list(new_attendance)
        print(f"Succesfully wrote {num} students to {self.output_path}!")


def find_file_path(file_str: str) -> str:
    current_directory = os.path.dirname(os.path.realpath(__file__))

    # Iterate over files in the script's directory
    for filename in os.listdir(current_directory):
        # Check if the filename starts with the constant string
        if filename.startswith(file_str):
            # Print the first matching filename and exit the loop
            print(f"Found file: {filename}!")
            return os.path.join(current_directory, filename)

    raise FileNotFoundError(errno.ENOENT, os.strerror(errno.ENOENT), file_str)


def main():
    try:
        tut_path = find_file_path(TUTORIAL_LIST_FILENAME)
        bot_name_path = find_file_path(ATTENDANCE_NAMES_FILE)
    except FileNotFoundError as e:
        print(e)
        return

    attendance = TakeAttendance(
        output_path=tut_path,
        input_path=bot_name_path,
        tutorial_number=TUTORIAL_NUMBER,
        append_mode=APPEND_MODE,
    )


if __name__ == "__main__":
    main()
