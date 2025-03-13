persian_numbers = {
    0: "صفر", 1: "یک", 2: "دو", 3: "سه", 4: "چهار", 5: "پنج", 6: "شش", 7: "هفت", 8: "هشت", 9: "نه",
    10: "ده", 11: "یازده", 12: "دوازده", 13: "سیزده", 14: "چهارده", 15: "پانزده", 16: "شانزده", 17: "هفده", 18: "هجده", 19: "نوزده",
    20: "بیست", 30: "سی", 40: "چهل", 50: "پنجاه", 60: "شصت", 70: "هفتاد", 80: "هشتاد", 90: "نود",
    100: "صد", 200: "دویست", 300: "سیصد", 400: "چهارصد", 500: "پانصد", 600: "ششصد", 700: "هفتصد", 800: "هشتصد", 900: "نهصد"
}

powers_of_thousand = [
    (1000000000, "میلیارد"),  # Billion
    (1000000, "میلیون"),      # Million
    (1000, "هزار"),           # Thousand
]

def convert_hundreds(number):
    """ Convert numbers from 1 to 999 to Persian words """
    if number < 20:
        return persian_numbers[number]
    elif number < 100:
        tens = number // 10 * 10
        remainder = number % 10
        return persian_numbers[tens] + ("" if remainder == 0 else " و " + persian_numbers[remainder])
    else:
        hundreds = number // 100 * 100
        remainder = number % 100
        return persian_numbers[hundreds] + ("" if remainder == 0 else " و " + convert_hundreds(remainder))

def convert_to_words(number):
    """ Convert any number to Persian words """
    if number == 0:
        return persian_numbers[0]

    words = []

    # Iterate over powers of thousand
    for value, word in powers_of_thousand:
        if number >= value:
            count = number // value
            words.append(convert_hundreds(count) + " " + word)
            number %= value

    if number > 0:
        words.append(convert_hundreds(number))

    return " و ".join(words)

# # Example usage:
# number = 3110000  # Three million, one hundred ten thousand
# persian_words = convert_to_words(number)
# print(persian_words)  # Outputs: "سه میلیون و صد و ده هزار"
