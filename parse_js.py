from py_mini_racer import MiniRacer

with open('static/resumo.js', 'r', encoding='utf-8') as f:
    js_code = f.read()

# Mock DOM objects so it doesn't crash on document.addEventListener during execution, if it executes.
# Actually, we just want to compile it.
ctx = MiniRacer()
try:
    ctx.eval(js_code)
    print("No syntax errors! (It evaluated without throwing syntax error)")
except Exception as e:
    print("Error:", e)
