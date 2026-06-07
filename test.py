import psycopg2

try:
    print("Executing check script to see if psycopg2 IN %s syntax generates valid SQL...")
    import psycopg2.extensions as ext
    cur = ext.adapt((("LTR", "DS"),)) # Adapt a tuple of strings
    print(cur.getquoted())
except Exception as e:
    print("Error:", e)
