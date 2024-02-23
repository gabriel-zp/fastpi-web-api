@app.post("/api/v1/endpoint2")
async def endpoint2(n: int, arquivo_bytes: UploadFile = File(...)):

    df = pd.read_excel(io.BytesIO(await arquivo_bytes.read()))

    if n > len(df.columns):
        return{"error": "Número de colunas inválido"}
    
    df = df.iloc[:, :-n]

    timestamp = datetime.now().strftime("%Y%m%d%H")

    outputFileName = f"temp/PlanilhaTeste_{timestamp}.xlsx"

    with pd.ExcelWriter(outputFileName, engine='openpyxl') as writer:
        df.to_excel(writer, index=False)
    
    print(outputFileName)
    return{"aaa": outputFileName}
