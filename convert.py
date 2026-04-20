import re

match = re.match(r"(\d{1,2}/\d{1,2})\s*(.*)", line)
if match:
    date = match.group(1)
    text = match.group(2)
else:
    date = ""
    text = line

timeline.append({
    "date": date,
    "text": text
})
