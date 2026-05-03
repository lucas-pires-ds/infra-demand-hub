import pyodbc

# Obtém a lista de drivers instalados
drivers = pyodbc.drivers()

# Imprime cada driver encontrado
print("Drivers SQL (ODBC) instalados:")
for driver in drivers:
    print(f"- {driver}")

    
