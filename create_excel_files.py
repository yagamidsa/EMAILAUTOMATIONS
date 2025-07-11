import pandas as pd
import os

print("🚀 Creando archivos Excel de ejemplo...")

# Crear carpetas si no existen
os.makedirs('data', exist_ok=True)
os.makedirs('adjuntos', exist_ok=True)
os.makedirs('reportes', exist_ok=True)

# 1. CAMPAÑAS.xlsx
print("📊 Creando CAMPAÑAS.xlsx...")
campanas_data = {
    'ID': [1, 2, 3],
    'Nombre_Campaña': [
        'Productos Julio',
        'Halloween 2025', 
        'Navidad 2025'
    ],
    'Asunto_Email': [
        'Boba Pearls plus Delicious Fudge and Seasonal Comforts from Sage Wholesale',
        '🎃 Spooky Halloween Inventory - Stock Up Now!',
        '🎅 Christmas Collection - Order Before Nov 15!'
    ],
    'Contenido_Email': [
        """Hello {NOMBRE},

{MENSAJE_PERSONAL}

Check out these items for some impulse opportunity now plus sure to come sales as we get into the 2nd half of the year!

🍹Boba Pearls are really cool, ever heard of them? They can be added to any beverage for a unique taste and satisfying texture.

🍫Gourmet Fudge, the kind you can get on the boardwalk, is available for your store this fall and holiday season.

Do you want to place an order? Please do not hesitate to ask. We hope to earn your business!

Enjoy your day,
{REMITENTE_NOMBRE}""",
        
        """Hello {NOMBRE},

Fall is here and Halloween is coming fast! We have some spooky good deals for {EMPRESA}.

🎃 Halloween decorations and candy are flying off the shelves.

Best regards,
{REMITENTE_NOMBRE}""",

        """Hello {NOMBRE},

The holiday season is approaching fast! Time to stock up on Christmas essentials for {EMPRESA}.

🎄 Our Christmas collection includes everything your customers need.

Happy Holidays,
{REMITENTE_NOMBRE}"""
    ],
    'ACTIVA': ['SÍ', 'NO', 'NO']
}

df_campanas = pd.DataFrame(campanas_data)
df_campanas.to_excel('data/CAMPAÑAS.xlsx', sheet_name='Campañas', index=False)
print("   ✅ CAMPAÑAS.xlsx creado")

# 2. CLIENTES.xlsx  
print("👥 Creando CLIENTES.xlsx...")
clientes_data = {
    'Email': [
        'test1@ejemplo.com',
        'test2@ejemplo.com',
        'test3@ejemplo.com', 
        'admin@prueba.com'
    ],
    'Nombre': [
        'Juan Carlos',
        'Maria Rodriguez',
        '',
        'Admin'
    ],
    'Empresa': [
        'MiniMarket Plus',
        'Quick Store',
        'Convenience Co',
        'Test Store'
    ],
    'Mensaje_Personal': [
        "Hope business is going well this month!",
        "Thanks for being a valued customer!",
        "Hope your summer sales are strong!",
        "Looking forward to working with you!"
    ]
}

df_clientes = pd.DataFrame(clientes_data)
df_clientes.to_excel('data/CLIENTES.xlsx', sheet_name='Contactos', index=False)
print("   ✅ CLIENTES.xlsx creado")

# 3. CONFIGURACION.xlsx
print("⚙️ Creando CONFIGURACION.xlsx...")
config_data = {
    'Configuración': [
        'Tu_Email',
        'Tu_Nombre', 
        'Tu_Empresa',
        'Total_Correos_Por_Dia',
        'Horas_Para_Enviar_Todo',
        'Correos_Por_Lote',
        'Minutos_Entre_Lotes',
        'Empezar_Inmediatamente'
    ],
    'Valor': [
        'sage-gefticosalessupport@geftico.com',
        'Admin',
        'Sage Wholesale',
        400,
        8,
        5,
        6,
        'SÍ'
    ]
}

df_config = pd.DataFrame(config_data)
df_config.to_excel('data/CONFIGURACION.xlsx', sheet_name='Config', index=False)
print("   ✅ CONFIGURACION.xlsx creado")

print("\n🎉 ¡Archivos Excel creados exitosamente!")
print("📁 Estructura final:")
print("   ├── data/CAMPAÑAS.xlsx")
print("   ├── data/CLIENTES.xlsx") 
print("   ├── data/CONFIGURACION.xlsx")
print("   ├── adjuntos/ (Vacía - aquí pondrás tus archivos)")
print("   └── reportes/ (Vacía - aquí se generarán reportes)")