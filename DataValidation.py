import datetime
import re

class DataValidation:
    # Validates user input against any criteria for numbers.
    # Set relevant parameters to False when calling this method to disallow those inputs (e.g., if you do NOT want to allow negative numbers, set negative_num = False). By default all rational numbers are allowed (all parameters are set to True).
    # If allow_back/allow_exit is True, user may enter "back"/"exit" (not case-sensitive).
    # min_num/max_num should only be set to integers or floats. If set to True, no effect.
    # Returns False is failed. Returns the user's input as string if succeeds.
    def validate_user_input_num(self, user_input, float_num = True, negative_num = True, zero_num = True, positive_num = True, min_num = False, max_num = False, allow_back = False, allow_exit = False):
        self.check_types_to_raise_exc(
            (user_input, float_num, negative_num, zero_num, positive_num, min_num, max_num, allow_back, allow_exit),
            ((str, int, float), bool, bool, bool, bool, (bool, int, float), (bool, int, float), bool, bool),
            ("user_input", "float_num", "negative_num", "zero_num", "positive_num", "min_num", "max_num", "allow_back", "allow_exit")
        )

        # Checking if all criteria are set to False (nothing would pass this check)
        if negative_num == False and zero_num == False and positive_num == False:
            raise InvalidValidateNumSettings

        # Invalid entry messages to user
        invalid_num = "\nYour input must be a number.\n\n(Press Enter.)\n"
        invalid_float_num = "\nYour input may not be a float. (Your input must be an integer).\n\n(Press Enter.)\n"
        invalid_negative_num = "\nYour input may not be negative.\n\n(Press Enter.)\n"
        invalid_zero_num = "\nYour input may not be zero.\n\n(Press Enter.)\n"
        invalid_positive_num = "\nYour input may not be positive.\n\n(Press Enter.)\n"
        invalid_min_num = f"\nYour input may not be below {min_num}."
        invalid_max_num = f"\nYour input may not be above {max_num}."

        user_input = user_input.strip().lower()
        orig_user_input = user_input.strip()

        if user_input == "": return False

        if allow_back == True:
            if user_input == "back": return "back"

        if allow_exit == True:
            if user_input == "exit": return "exit"

        # Checking that user's input is a number
        try: user_input = float(user_input)
        except ValueError:
            input(invalid_num)
            return False

        if float_num == False:
            if user_input != int(user_input):
                input(invalid_float_num)
                return False

        if negative_num == False:
            if user_input < 0:
                input(invalid_negative_num)
                return False

        if zero_num == False:
            if user_input == 0:
                input(invalid_zero_num)
                return False

        if positive_num == False:
            if user_input > 0:
                input(invalid_positive_num)
                return False

        if min_num != False:
            if user_input < min_num:
                input(invalid_min_num)
                return False

        if max_num != False:
            if user_input > max_num:
                input(invalid_max_num)
                return False

        return orig_user_input

    # Validates user input based on custom tuple in "acceptable" argument.
    # This does not loop; this should be called within a loop obtaining user's input.
    # Not case-senstiive. E.g., "mn" matches "MN".
    # If allow_back is True, user may enter "back".
    # Returns False for a failed check. Returns "back" if user wants to go back. Returns stripped user's input if succeeded.
    def validate_user_input_custom(self, user_input, acceptable, allow_back = False, allow_exit = False):
        self.check_types_to_raise_exc((user_input, acceptable, allow_back), (str, (list, tuple), bool), ("user_input", "acceptable", "allow_back"))

        user_input = user_input.strip()
        user_input_l = user_input.lower()

        if user_input == "": return False

        if allow_back == True:
            if user_input_l == "back": return "back"

        if allow_exit == True:
            if user_input_l == "exit": return "exit"

        if user_input in acceptable or user_input.capitalize() in acceptable or user_input_l in acceptable or user_input.upper() in acceptable: return user_input

        return False

    # Validates user input as a date.
    # This does not loop; this should be called within a loop obtaining user's input.
    # NOT case-sensitive. Commas do not matter.
    # If allow_back is True, user may enter "back".
    # Returns False for a failed check. Returns "back" if user wants to go back. Returns datetime object if succeeded.
    def validate_user_input_date(self, user_input, allow_back = False, allow_exit = False):
        self.check_types_to_raise_exc((user_input, allow_back), (str, bool), ("user_input", "allow_back"))

        if user_input == "": return False

        user_input = user_input.strip().lower().replace(",", "").replace("-", "/").replace(".", "/")
        user_input_l = user_input.lower()

        if allow_back == True:
            if user_input_l == "back": return "back"

        if allow_exit == True:
            if user_input_l == "exit": return "exit"

        # Tuples to be cycled through for checks
        # no_year_input_test is testing when the user enters a month and day and the year will then be assumed to be the current year.
        no_year_input_test = (
            ("/" + str(datetime.datetime.now().year), "%m/%d/%Y"),
            (" " + str(datetime.datetime.now().year), "%B %d %Y"),
            (" " + str(datetime.datetime.now().year), "%b %d %Y")
        )

        # Checks for instances when the user did not provide a year
        for check in no_year_input_test:
            try: time_obj = datetime.datetime.strptime(user_input + check[0], check[1])
            except: pass
            else: return time_obj

        # input_test is testing when the user enters a date including the year
        # [0] is what strptime will use. [1] is True if validation was successful with "%y", meaning user provided only two digits for the year, not 4; in this case, any %y in the next 10 years will be assumed to be part of this century and anything else last (e.g., if now().year == 2022, then 1/1/32 is 1/1/2032 but 1/1/33 is 1/1/1933).
        input_test = (
            ("%m/%d/%Y", False),
            ("%m/%d/%y", True),
            ("%b %d %Y", False),
            ("%b %d %y", True),
            ("%B %d %Y", False),
            ("%B %d %y", True),
            ("%Y/%m/%d", False)
        )

        # Checks for instances when the user provided the full date
        for check in input_test:
            try: time_obj = datetime.datetime.strptime(user_input, check[0])
            except: pass
            else:
                # If validation succeeded with "%y", then checks to see if current or last century should be used.
                if check[1] and \
                        time_obj.year - datetime.datetime.now().year > 10:
                    time_obj = datetime.datetime(time_obj.year - 100, time_obj.month, time_obj.day)

                return time_obj

        return False

    # Validates user_input as an email address.
    # Checks general formatting with regex and then checks for invalid characters.
    def validate_user_input_email(self, user_input):
        invalid_chars = (":", "–", "[", "´", ";", "—", "]", "&", ",", "<", "(", "\\", "“", '"', ">", ")", " ")
        user_input = user_input.strip()

        if re.fullmatch(".+@.+\..+", user_input) == None: return False
        if user_input.count("@") != 1: return False

        for char in invalid_chars:
            if char in user_input: return False

        return user_input

    # Validates user input against a regular expression. By default, searches user_input for regular expression. If fullmatch == True, entire user_input must match.
    def validate_user_input_regex(self, user_input, regex, fullmatch = False, allow_back = False, allow_exit = False):
        user_input = user_input.strip()
        user_input_l = user_input.lower()

        if allow_exit == True:
            if user_input_l == "exit": return "exit"

        if allow_back == True:
            if user_input_l == "back": return "back"

        if fullmatch == True:
            if re.fullmatch(regex, user_input) == None: return False
        else:
            if re.search(regex, user_input) == None: return False

        return user_input

    # Checks numerous variables to ensure they are the correct type. Raises exception if type is incorrect.
    # All arguments MUST be lists/tuples, even if they have only one element. (Note that if checking just one element, just doing the check directly, without check_to_raise_exc(), and then directly callin InvalidTypePassed(), is better.)
    # vars_to_check is a list/tuple of all variables to validate type
    # types_to_compare is a list/tuple of all types (must use type here, not string)
    # vars_as_strings is a list/tuple of strings of all variables being evaluated. This list visually is identical to vars_to_check except each element is a string (is in quotation marks).
    def check_types_to_raise_exc(self, vars_to_check, types_to_compare, vars_as_strings):
        # Before validating the data types provided, method first validates that the lists/tuples are the same length and that they are in fact lists or tuples.
        # Checks that the list lengths match (the lists will be zipped)
        if len(vars_to_check) != len(types_to_compare) or \
                len(types_to_compare) != len(vars_as_strings): raise InvalidListLength((vars_to_check, types_to_compare, vars_as_strings))

        # Checks that the arguments are lists or tuples
        validate_vars = zip((vars_to_check, types_to_compare, vars_as_strings), ("vars_to_check", "types_to_compare", "vars_as_strings"))
        for tup in validate_vars:
            if isinstance(tup[0], (list, tuple)) == False: raise InvalidTypePassed(tup[1], type(vars_to_check), (list, tuple))

        # Checks that the variables (vars_to_check) match the types provided (types_to_compare). Failure raises an Exception and informs the user of the problematic variable.
        list_to_check = zip(vars_to_check, types_to_compare, vars_as_strings)
        for checks in list_to_check:
            if isinstance(checks[0], checks[1]) == False: raise InvalidTypePassed(checks[2], type(checks[0]), checks[1])

# ----------------------------EXCEPTIONS CLASSES----------------------------
# This exception is available for any method to check a variable type. An invalid type will raise this error.
# To check multiple variables at once, use WedriverMain() method check_types_to_raise_exc(). That method loops and checks each variable with the below class.
# relevant_variable is a string that can be printed to the user to identify which variable is invalid
# type_passed is the actual type of the variable (valid or invalid). Use type() around the variable i question for this.
# type_needed is the type itself that the code requires (e.g., just type "str" or "float" but without quotation marks)
class InvalidTypePassed(Exception):
    def __init__(self, relevant_variable, type_passed, type_needed):
        message = f"Argument {relevant_variable} must be {type_needed}. Received {type_passed}."
        super().__init__(message)

# Exception if lists/tuples provided to a method have unmatched lengths when they must match (e.g., they will be zipped).
class InvalidListLength(Exception):
    def __init__(self, lists_tuples):
        lists_tuples_to_user = ", ".join([l_t for l_t in lists_tuples])
        message = f"These lists/tuples have unmatched lengths: {lists_tuples_to_user}. Lengths must match."
        super().__init__(message)

class InvalidValidateNumSettings(Exception):
    def __init__(self):
        message = "All real numbers are excluded by this criteria. Modify code to allow some user input to pass."
        super().__init__(message)