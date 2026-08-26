import sqlite3
conn = sqlite3.connect('banco_resumo.db')
conn.execute("UPDATE configuracoes SET valor = 'gemini-3.1-flash-lite' WHERE chave = 'gemini_model'")
conn.commit()
conn.close()
print("Corrigido")
