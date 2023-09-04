def get_hex(ch):
    fch = ""
    for j in ch:
        if fch == "":
            fch+=f"ChrW({ord(j)})"
        else:
            fch+=f" & ChrW({ord(j)})"
    return fch



numbers = [
    "शून्य", "एक", "दुई", "तिन", "चार", "पाँच", "छ", "सात", "आठ", "नौँ","दश",
    "एघार", "बाह्र", "तेह्र", "चौध","पन्द्र","सोह्र","सत्र","अठार","उन्नाइस","बिस",
    "एक्काइस","बाइस","तेइस","चौबिस","पच्चिस","छब्बिस","सत्ताइस","अठ्ठाइस","उनन्तिस","तिस",
    "एकतिस","बत्तिस","तेत्तिस","चौतिस","पैतिस","छत्तिस","सड्तिस","अड्तिस","उन्चालिस","चालिस"
    "एकचालिस","बयालिस","तिर्चालिस","चौवालिस","पैतालिस","छयालिस","सड्चालिस","अड्चालिस","उनन्पचास","पचास",
    "एकाउन्न","बाउन्न","तिर्पन्न","चौवन्न","पच्पन्न","छपन्न","सन्ताउन्न","अन्ठाउन्न","उन्साठ्ठि","साठ्ठि",
    "एकसठ्ठि","बैसठ्ठि","तिर्सठ्ठि","चौसठ्ठि","पैसठ्ठि","छैसठ्ठि","सड्सठ्ठि","अडसठ्ठि","उन्सत्तरी","सत्तरी",
    "एकतर","बहतर","तिर्हतर","चौरत्तर","पचत्तर","छयत्तर","सत्ततर","अठत्तर","उन्नासी","असि",
    "एकासी","बयासी","तिरासी","चौरासी","पचासी","छयासी","सतासी","अठासी","उननब्बे","नब्बे",
    "एकानब्बे","बयानब्बे","तिरानब्बे","चौरानब्बे","पञ्चानब्बे","छयानब्बे","सन्तानब्बे","अन्ठानब्बे","उनन्सय"
]

specials = [
    " सय "," हजार ", " लाख "," करोड "," अरब "," खरब "
]

total_length = sum([len(get_hex(i)) for i in numbers])

with open("NepaliNumerics.bas", "w") as f:
    with open("first_part.txt", "r") as fp:
        text = fp.read()
    f.write(text)
    text = ""
    length = 0
    previous = 0

    for count, num in enumerate(numbers):
        length += len(get_hex(num))
        if length > 700:
            f.write(", ".join(get_hex(i) for i in numbers[previous:count]))
            f.write(", _\n\t\t\t\t\t\t")

            previous = count
            length = 0

    f.write(", ".join(get_hex(i) for i in numbers[previous:count+1]))
    f.write(")\n\n\tSpecials = Array(")
    f.write(", ".join(get_hex(i) for i in specials))
    f.write(")")

    with open("second_part.txt", "r") as sp:
        text = sp.read()
    f.write(text)