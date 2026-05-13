"""Simulate VBA JVal against live Oracle JSON to find Category parsing issue."""
import requests
from requests.auth import HTTPBasicAuth

BASE = "https://eese.fa.us8.oraclecloud.com/fscmRestApi/resources/11.13.18.05"
AUTH = HTTPBasicAuth("Kevin.a.Budziszewski", "GaryTheCat2929!")

r = requests.get(BASE + "/invoices/300000046385743/child/attachments", auth=AUTH)
raw = r.text

# Show the exact bytes around "Category"
idx = raw.find('"Category"')
print("Category context:", repr(raw[idx : idx + 60]))
print()

def jval(json_str, key):
    """Python port of the VBA JVal function."""
    pat = '"' + key + '"'
    sp = json_str.find(pat)
    if sp == -1:
        return "<<KEY NOT FOUND>>"
    sp += len(pat)
    # skip spaces before colon
    while sp < len(json_str) and json_str[sp] == " ":
        sp += 1
    if sp >= len(json_str) or json_str[sp] != ":":
        return "<<NO COLON, got: " + repr(json_str[sp:sp+5]) + ">>"
    sp += 1
    # skip spaces after colon
    while sp < len(json_str) and json_str[sp] == " ":
        sp += 1
    if json_str[sp] == '"':
        sp += 1
        ep = sp
        while ep < len(json_str):
            if json_str[ep] == '"' and json_str[ep - 1] != '\\':
                break
            ep += 1
        return json_str[sp:ep]
    elif json_str[sp:sp+4] == "null":
        return ""
    else:
        e1 = json_str.find(",", sp)
        e2 = json_str.find("}", sp)
        if e1 == -1: e1 = len(json_str)
        if e2 == -1: e2 = len(json_str)
        return json_str[sp : min(e1, e2)].strip()

# Simulate ParseItems
items_start = raw.find('"items"')
bracket = raw.find("[", items_start) + 1
depth, ostart, items = 0, 0, []
p = bracket
while p < len(raw):
    c = raw[p]
    if c == "{":
        if depth == 0:
            ostart = p
        depth += 1
    elif c == "}":
        depth -= 1
        if depth == 0 and ostart:
            items.append(raw[ostart : p + 1])
            ostart = 0
    elif c == "]" and depth == 0:
        break
    p += 1

print(f"Items found by ParseItems: {len(items)}")
for i, item in enumerate(items):
    cat   = jval(item, "Category")
    fname = jval(item, "FileName")
    fhref = None
    # simulate FileContentsHref
    search_from = 0
    while True:
        np = item.find('"name"', search_from)
        if np == -1:
            break
        cp = item.find(":", np + 6) + 1
        while cp < len(item) and item[cp] == " ":
            cp += 1
        if item[cp:cp+14] == '"FileContents"':
            hp = item.rfind('"href"', 0, np)
            if hp != -1:
                hc = item.find(":", hp + 6) + 1
                while hc < len(item) and item[hc] == " ":
                    hc += 1
                if item[hc] == '"':
                    hc += 1
                    he = item.find('"', hc)
                    fhref = item[hc:he]
            break
        search_from = np + 1

    print(f"  Item {i}: FileName={fname!r}")
    print(f"           Category={cat!r}")
    print(f"           FileContentsHref={'<found>' if fhref else '<NOT FOUND>'}")
    if fhref:
        print(f"           href={fhref[:80]}...")
    print()
