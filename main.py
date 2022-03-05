import pandas as pd
import win32com.client as win32


if __name__ == "__main__":
    # importar base de dados
    table = pd.read_excel("Vendas.xlsx")
    # visualizar base de dados
    pd.set_option("display.max_columns", None)
    # faturamento por loja
    with open("faturamento.txt", "w", encoding="utf8") as file:
        file.writelines(
            str(table[["ID Loja", "Valor Final"]].groupby("ID Loja").sum()))
    # quantidade de produtos vendidos por loja
    with open("quantidade_venda.txt", "w", encoding="utf8") as file:
        file.writelines(
            str(table[["ID Loja", "Quantidade"]].groupby("ID Loja").sum()))
    # ticket medio -> produto -> loja
    total_prices = table[["ID Loja", "Valor Final"]].groupby("ID Loja").sum()
    quantity = table[["ID Loja", "Quantidade"]].groupby("ID Loja").sum()
    ticket_media = (total_prices["Valor Final"] /
                    quantity["Quantidade"]).to_frame()
    ticket_media = ticket_media.rename(columns={0: "Ticket Médio"})
    # enviar e-mail com relatório
    outlook = win32.Dispatch("outlook.application")
    mail = outlook.CreateItem(0)
    mail.To = "theodorojoaogabriel@gmail.com"
    mail.Subject = "Dados em Python - empresa - By João Theodoro"
    mail.HTMLBody = f'''
        <h1>Relatório de vendas por loja</h1>

        <h2>Faturamento</h2>
        {total_prices.to_html(formatters={"Valor Final": "R${:,.2f}".format})}
        <h2>Faturamento</h2>
        {quantity.to_html()}
        <h2>Ticket médio de produto em cada loja</h2>
        {ticket_media.to_html(formatters={"Ticket Médio": "R${:,.2f}".format})}
        <h3><strong>Duvidas para: joaogabrielvtheodoro@hotmail.com</strong></h3>
    '''

    mail.Send()
