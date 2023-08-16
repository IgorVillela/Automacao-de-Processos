import smtplib
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from unidecode import unidecode


class EnvioEmail:

    def __init__(self):
        self._senha_app = "SENHA_APP"  # SENHA APP GMAIL

    def enviar_email(self, destinatario, gerente, loja, data,
                     faturamento_dia, valor_meta_faturamento_dia, ticket_dia, valor_meta_ticket_dia,
                     qtde_produtos_diversos, valor_meta_diversidade,
                     faturamento_ano, valor_meta_faturamento_ano, ticket_ano, valor_meta_ticket_ano,
                     qtde_produtos_diversos_ano, valor_meta_diversidade_ano,
                     caminho_arquivo1, caminho_arquivo2):
        sender_email = "E-MAIL REMETENTE" # E-MAIL REMETENTE
        sender_password = self._senha_app  # SENHA APP GMAIL
        recipient_email = destinatario # E-MAIL DESTINARÁRIO
        subject = f"OnePage Dia {data} - Loja {loja}"
        body = f""" <p style='margin-top:0cm;margin-right:0cm;margin-bottom:8.0pt;margin-left:0cm;font-size:11.0pt
        ;font-family:"Calibri",sans-serif;'>Bom dia, {gerente}</p> <p style='margin-top:0cm;margin-right:0cm;margin
        -bottom:8.0pt;margin-left:0cm;font-size:11.0pt;font-family:"Calibri",sans-serif;'>&nbsp;</p> <p 
        style='margin-top:0cm;margin-right:0cm;margin-bottom:8.0pt;margin-left:0cm;font-size:11.0pt;font-family
        :"Calibri",sans-serif;'>O resultado de ontem ({data}) da Loja {loja} foi:</p> <table 
        style="border-collapse:collapse;border:none;"> <tbody> <tr> <td style="width: 106.15pt;padding: 0cm 
        5.4pt;vertical-align: top;"> <p style='margin-top:0cm;margin-right:0cm;margin-bottom:0cm;margin-left:0cm;font
        -size:11.0pt;font-family:"Calibri",sans-serif;line-height:normal;'><span 
        style="font-size:16px;">&nbsp;</span></p> </td> <td style="width: 106.15pt;padding: 0cm 5.4pt;vertical-align: 
        top;"> <p style='margin-top:0cm;margin-right:0cm;margin-bottom:0cm;margin-left:0cm;font-size:11.0pt;font
        -family:"Calibri",sans-serif;text-align:center;line-height:normal;'><span style="font-size:16px;">Valor 
        Dia</span></p> </td> <td style="width: 106.2pt;padding: 0cm 5.4pt;vertical-align: top;"> <p 
        style='margin-top:0cm;margin-right:0cm;margin-bottom:0cm;margin-left:0cm;font-size:11.0pt;font-family
        :"Calibri",sans-serif;text-align:center;line-height:normal;'><span style="font-size:16px;">Meta 
        Dia</span></p> </td> <td style="width: 106.2pt;padding: 0cm 5.4pt;vertical-align: top;"> <p 
        style='margin-top:0cm;margin-right:0cm;margin-bottom:0cm;margin-left:0cm;font-size:11.0pt;font-family
        :"Calibri",sans-serif;text-align:center;line-height:normal;'><span style="font-size:16px;">Cen&aacute;rio 
        Dia</span></p> </td> </tr> <tr> <td style="width: 106.15pt;padding: 0cm 5.4pt;vertical-align: top;"> <p 
        style='margin-top:0cm;margin-right:0cm;margin-bottom:0cm;margin-left:0cm;font-size:11.0pt;font-family
        :"Calibri",sans-serif;text-align:right;line-height:normal;'><span 
        style="font-size:16px;">Faturamento</span></p> </td> <td style="width: 106.15pt;padding: 0cm 
        5.4pt;vertical-align: top;"> <p style='margin-top:0cm;margin-right:0cm;margin-bottom:0cm;margin-left:0cm;font
        -size:11.0pt;font-family:"Calibri",sans-serif;text-align:center;line-height:normal;'>R$ 
        {faturamento_dia:,.2f}</p> </td> <td style="width: 106.2pt;padding: 0cm 5.4pt;vertical-align: top;"> <p 
        style='margin-top:0cm;margin-right:0cm;margin-bottom:0cm;margin-left:0cm;font-size:11.0pt;font-family:"Calibri",
        sans-serif;text-align:center;line-height:normal;'>R$ {valor_meta_faturamento_dia:,.2f}</p> </td> <td 
        style="width: 106.2pt;padding: 0cm 5.4pt;vertical-align: top;"> <p 
        style='margin-top:0cm;margin-right:0cm;margin-bottom:0cm;margin-left:0cm;font-size:11.0pt;font-family 
        :"Calibri", sans-serif;text-align:center;line-height:normal;'><span 
        style='font-size:16px;font-family:"Arial",sans-serif;color: 
{'green' if faturamento_dia >= valor_meta_faturamento_dia else 'red'};'>◙</span></p> </td> </tr> <tr> <td 
style="width: 106.15pt;padding: 0cm 5.4pt;vertical-align: top;"> <p 
style='margin-top:0cm;margin-right:0cm;margin-bottom:0cm;margin-left:0cm;font-size:11.0pt;font-family:"Calibri", 
sans-serif;text-align:right;line-height:normal;'><span style="font-size:16px;">Ticket M&eacute;dio</span></p> </td> 
<td style="width: 106.15pt;padding: 0cm 5.4pt;vertical-align: top;"> <p 
style='margin-top:0cm;margin-right:0cm;margin-bottom:0cm;margin-left:0cm;font-size:11.0pt;font-family:"Calibri", 
sans-serif;text-align:center;line-height:normal;'>R$ {ticket_dia:,.2f}</p> </td> <td style="width: 106.2pt;padding: 
0cm 5.4pt;vertical-align: top;"> <p style='margin-top:0cm;margin-right:0cm;margin-bottom:0cm;margin-left:0cm;font 
-size:11.0pt;font-family:"Calibri",sans-serif;text-align:center;line-height:normal;'>R$ {valor_meta_ticket_dia:,.2f} 
</p> </td> <td style="width: 106.2pt;padding: 0cm 5.4pt;vertical-align: top;"> <p 
style='margin-top:0cm;margin-right:0cm;margin-bottom:0cm;margin-left:0cm;font-size:11.0pt;font-family:"Calibri", 
sans-serif;text-align:center;line-height:normal;'><span style='font-size:16px;font-family:"Arial",sans-serif;color: 
{'green' if ticket_dia >= valor_meta_ticket_dia else 'red'};'>◙</span></p> </td> </tr> <tr> <td style="width: 
106.15pt;padding: 0cm 5.4pt;vertical-align: top;"> <p 
style='margin-top:0cm;margin-right:0cm;margin-bottom:0cm;margin-left:0cm;font-size:11.0pt;font-family:"Calibri", 
sans-serif;text-align:right;line-height:normal;'><span style="font-size:16px;">Produtos distintos</span></p> </td> 
<td style="width: 106.15pt;padding: 0cm 5.4pt;vertical-align: top;"> <p 
style='margin-top:0cm;margin-right:0cm;margin-bottom:0cm;margin-left:0cm;font-size:11.0pt;font-family:"Calibri", 
sans-serif;text-align:center;line-height:normal;'>{qtde_produtos_diversos}</p> </td> <td style="width:
        106.2pt;padding: 0cm 5.4pt;vertical-align: top;"> <p 
        style='margin-top:0cm;margin-right:0cm;margin-bottom:0cm;margin-left:0cm;font-size:11.0pt;font-family:"Calibri",
        sans-serif;text-align:center;line-height:normal;'>{valor_meta_diversidade}</p> </td> <td style="width: 
        106.2pt;padding: 0cm 5.4pt;vertical-align: top;"> <p 
        style='margin-top:0cm;margin-right:0cm;margin-bottom:0cm;margin-left:0cm;font-size:11.0pt;font-family
        :"Calibri", sans-serif;text-align:center;line-height:normal;'><span 
        style='font-size:16px;font-family:"Arial",sans-serif;color: 
{'green' if qtde_produtos_diversos >= valor_meta_diversidade else 'red'};'>◙</span></p> </td> </tr> </tbody> </table> 
<p style='margin-top:0cm;margin-right:0cm;margin-bottom:8.0pt;margin-left:0cm;font-size:11.0pt;font-family:"Calibri", 
sans-serif;'>&nbsp;</p> <table style="border-collapse:collapse;border:none;"> <tbody> <tr> <td style="width: 
106.15pt;padding: 0cm 5.4pt;vertical-align: top;"> <p 
style='margin-top:0cm;margin-right:0cm;margin-bottom:0cm;margin-left:0cm;font-size:11.0pt;font-family:"Calibri", 
sans-serif;line-height:normal;'><span style="font-size:16px;">&nbsp;</span></p> </td> <td style="width: 
106.15pt;padding: 0cm 5.4pt;vertical-align: top;"> <p 
style='margin-top:0cm;margin-right:0cm;margin-bottom:0cm;margin-left:0cm;font-size:11.0pt;font-family:"Calibri", 
sans-serif;text-align:center;line-height:normal;'><span style="font-size:16px;">Valor Ano</span></p> </td> <td 
style="width: 106.2pt;padding: 0cm 5.4pt;vertical-align: top;"> <p 
style='margin-top:0cm;margin-right:0cm;margin-bottom:0cm;margin-left:0cm;font-size:11.0pt;font-family:"Calibri", 
sans-serif;text-align:center;line-height:normal;'><span style="font-size:16px;">Meta Ano</span></p> </td> <td 
style="width: 106.2pt;padding: 0cm 5.4pt;vertical-align: top;"> <p 
style='margin-top:0cm;margin-right:0cm;margin-bottom:0cm;margin-left:0cm;font-size:11.0pt;font-family:"Calibri", 
sans-serif;text-align:center;line-height:normal;'><span style="font-size:16px;">Cen&aacute;rio Ano</span></p> </td> 
</tr> <tr> <td style="width: 106.15pt;padding: 0cm 5.4pt;vertical-align: top;"> <p 
style='margin-top:0cm;margin-right:0cm;margin-bottom:0cm;margin-left:0cm;font-size:11.0pt;font-family:"Calibri", 
sans-serif;text-align:right;line-height:normal;'><span style="font-size:16px;">Faturamento</span></p> </td> <td 
style="width: 106.15pt;padding: 0cm 5.4pt;vertical-align: top;"> <p 
style='margin-top:0cm;margin-right:0cm;margin-bottom:0cm;margin-left:0cm;font-size:11.0pt;font-family:"Calibri", 
sans-serif;text-align:center;line-height:normal;'>R$ {faturamento_ano:,.2f}</p> </td> <td style="width:
        106.2pt;padding: 0cm 5.4pt;vertical-align: top;"> <p 
        style='margin-top:0cm;margin-right:0cm;margin-bottom:0cm;margin-left:0cm;font-size:11.0pt;font-family:"Calibri",
        sans-serif;text-align:center;line-height:normal;'>R$ {valor_meta_faturamento_ano:,.2f}</p> </td> <td 
        style="width: 106.2pt;padding: 0cm 5.4pt;vertical-align: top;"> <p 
        style='margin-top:0cm;margin-right:0cm;margin-bottom:0cm;margin-left:0cm;font-size:11.0pt;font-family 
        :"Calibri", sans-serif;text-align:center;line-height:normal;'><span 
        style='font-size:16px;font-family:"Arial",sans-serif;color: 
{'green' if faturamento_ano >= valor_meta_faturamento_ano else 'red'};'>◙</span></p> </td> </tr> <tr> <td 
style="width: 106.15pt;padding: 0cm 5.4pt;vertical-align: top;"> <p 
style='margin-top:0cm;margin-right:0cm;margin-bottom:0cm;margin-left:0cm;font-size:11.0pt;font-family:"Calibri", 
sans-serif;text-align:right;line-height:normal;'><span style="font-size:16px;">Ticket M&eacute;dio</span></p> </td> 
<td style="width: 106.15pt;padding: 0cm 5.4pt;vertical-align: top;"> <p 
style='margin-top:0cm;margin-right:0cm;margin-bottom:0cm;margin-left:0cm;font-size:11.0pt;font-family:"Calibri", 
sans-serif;text-align:center;line-height:normal;'>R$ {ticket_ano:,.2f}</p> </td> <td style="width: 106.2pt;padding: 
0cm 5.4pt;vertical-align: top;"> <p style='margin-top:0cm;margin-right:0cm;margin-bottom:0cm;margin-left:0cm;font 
-size:11.0pt;font-family:"Calibri",sans-serif;text-align:center;line-height:normal;'>R$ 
{valor_meta_ticket_ano:,.2f}</p> </td> <td style="width: 106.2pt;padding: 0cm 5.4pt;vertical-align: top;"> <p 
style='margin-top:0cm;margin-right:0cm;margin-bottom:0cm;margin-left:0cm;font-size:11.0pt;font-family:"Calibri", 
sans-serif;text-align:center;line-height:normal;'><span style='font-size:16px;font-family:"Arial",sans-serif;color: 
{'green' if ticket_ano >= valor_meta_ticket_ano else 'red'};'>◙</span></p> </td> </tr> <tr> <td style="width: 
106.15pt;padding: 0cm 5.4pt;vertical-align: top;"> <p 
style='margin-top:0cm;margin-right:0cm;margin-bottom:0cm;margin-left:0cm;font-size:11.0pt;font-family:"Calibri", 
sans-serif;text-align:right;line-height:normal;'><span style="font-size:16px;">Produtos distintos</span></p> </td> 
<td style="width: 106.15pt;padding: 0cm 5.4pt;vertical-align: top;"> <p 
style='margin-top:0cm;margin-right:0cm;margin-bottom:0cm;margin-left:0cm;font-size:11.0pt;font-family:"Calibri", 
sans-serif;text-align:center;line-height:normal;'>{qtde_produtos_diversos_ano}</p> </td> <td style="width:
        106.2pt;padding: 0cm 5.4pt;vertical-align: top;"> <p 
        style='margin-top:0cm;margin-right:0cm;margin-bottom:0cm;margin-left:0cm;font-size:11.0pt;font-family:"Calibri",
        sans-serif;text-align:center;line-height:normal;'>{valor_meta_diversidade_ano}</p> </td> <td style="width: 
        106.2pt;padding: 0cm 5.4pt;vertical-align: top;"> <p 
        style='margin-top:0cm;margin-right:0cm;margin-bottom:0cm;margin-left:0cm;font-size:11.0pt;font-family
        :"Calibri", sans-serif;text-align:center;line-height:normal;'><span 
        style='font-size:16px;font-family:"Arial",sans-serif;color: 
{'green' if qtde_produtos_diversos_ano >= valor_meta_diversidade_ano else 'red'};'>◙</span></p> </td> </tr> </tbody> 
</table> <p style='margin-top:0cm;margin-right:0cm;margin-bottom:8.0pt;margin-left:0cm;font-size:11.0pt;font-family 
:"Calibri",sans-serif;'>&nbsp;</p> <p style='margin-top:0cm;margin-right:0cm;margin-bottom:8.0pt;margin-left:0cm;font 
-size:11.0pt;font-family:"Calibri",sans-serif;'>Segue em anexo a planilha com todos os dados para mais detalhes.</p> 
<p style='margin-top:0cm;margin-right:0cm;margin-bottom:8.0pt;margin-left:0cm;font-size:11.0pt;font-family:"Calibri", 
sans-serif;'>Qualquer d&uacute;vida, estou &agrave; disposi&ccedil;&atilde;o.</p> <p 
style='margin-top:0cm;margin-right:0cm;margin-bottom:8.0pt;margin-left:0cm;font-size:11.0pt;font-family:"Calibri", 
sans-serif;'>&nbsp;</p> <p style='margin-top:0cm;margin-right:0cm;margin-bottom:8.0pt;margin-left:0cm;font-size:11 
.0pt;font-family:"Calibri",sans-serif;'>Att.,</p> <p 
style='margin-top:0cm;margin-right:0cm;margin-bottom:8.0pt;margin-left:0cm;font-size:11.0pt;font-family:"Calibri", 
sans-serif;'>Igor</p>"""

        message = MIMEMultipart()
        message['Subject'] = subject
        message['From'] = sender_email
        message['To'] = recipient_email
        html_message = MIMEText(body, 'html')
        message.attach(html_message)

        part = MIMEBase('application', "octet-stream")
        part.set_payload(open(caminho_arquivo1, 'rb').read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', f'attachment; filename="{unidecode(loja)}{data}_VendasAnual.xlsx"')
        message.attach(part)

        part = MIMEBase('application', "octet-stream")
        part.set_payload(open(caminho_arquivo2, 'rb').read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', f'attachment; filename="{unidecode(loja)}{data}_VendasDia.xlsx"')
        message.attach(part)

        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as server:
            server.login(sender_email, sender_password)
            server.sendmail(sender_email, recipient_email, message.as_string())

        print('E-mail enviado!')

    def enviar_email_diretoria(self, destinatario, gerente, data, top_nome_loja_dia, top_faturamento_dia,
                               top_nome_loja_ano, top_faturamento_ano, pior_nome_loja_dia, pior_faturamento_dia,
                               pior_nome_loja_ano, pior_faturamento_ano, caminho_top_ano, caminho_pior_ano,
                               caminho_top_dia, caminho_pior_dia):
        sender_email = "igorvillela1@gmail.com"
        sender_password = self._senha_app  # SENHA APP GMAIL
        recipient_email = destinatario
        subject = f"Resumo Vendas {data}"
        body = f""" <p style='margin-top:0cm;margin-right:0cm;margin-bottom:8.0pt;margin-left:0cm;font-size:11.0pt
        ;font-family:"Calibri",sans-serif;'>Bom dia, {gerente}</p> <p style='margin-top:0cm;margin-right:0cm;margin
        -bottom:8.0pt;margin-left:0cm;font-size:11.0pt;font-family:"Calibri",sans-serif;'>&nbsp;</p> <p 
        style='margin-top:0cm;margin-right:0cm;margin-bottom:8.0pt;margin-left:0cm;font-size:11.0pt;font-family
        :"Calibri",sans-serif;'>A TOP loja com maior faturamento no dia foi a loja {top_nome_loja_dia} (R$ 
{top_faturamento_dia:,.2f})</p> <p style='margin-top:0cm;margin-right:0cm;margin-bottom:8.0pt;margin-left:0cm;font
-size:11.0pt;font-family:"Calibri",sans-serif;'>A TOP loja com maior faturamento no ano foi a loja 
{top_nome_loja_ano} (R$ {top_faturamento_ano:,.2f})</p> <p style='margin-top:0cm;margin-right:0cm;margin-bottom:8.0pt
;margin-left:0cm;font-size:11.0pt;font-family:"Calibri",sans-serif;'>A pior loja com menor faturamento no dia foi a 
loja {pior_nome_loja_dia} (R$ {pior_faturamento_dia:,.2f})</p> <p style='margin-top:0cm;margin-right:0cm;margin
-bottom:8.0pt;margin-left:0cm;font-size:11.0pt;font-family:"Calibri",sans-serif;'>A pior loja com menor faturamento 
no ano foi a loja {pior_nome_loja_ano} (R$ {pior_faturamento_ano:,.2f})</p> <p style='margin-top:0cm;margin-right:0cm
;margin-bottom:8.0pt;margin-left:0cm;font-size:11.0pt;font-family:"Calibri",sans-serif;'>Segue em anexo a planilha 
com todos os dados para mais detalhes.</p> <p style='margin-top:0cm;margin-right:0cm;margin-bottom:8.0pt;margin-left
:0cm;font-size:11.0pt;font-family:"Calibri",sans-serif;'>Qualquer d&uacute;vida, estou &agrave; 
disposi&ccedil;&atilde;o.</p> <p style='margin-top:0cm;margin-right:0cm;margin-bottom:8.0pt;margin-left:0cm;font-size
:11.0pt;font-family:"Calibri",sans-serif;'>&nbsp;</p> <p 
style='margin-top:0cm;margin-right:0cm;margin-bottom:8.0pt;margin-left:0cm;font-size:11.0pt;font-family:"Calibri",
sans-serif;'>Att.,</p> <p style='margin-top:0cm;margin-right:0cm;margin-bottom:8.0pt;margin-left:0cm;font-size:11.0pt
;font-family:"Calibri",sans-serif;'>Igor</p>"""
        message = MIMEMultipart()
        message['Subject'] = subject
        message['From'] = sender_email
        message['To'] = recipient_email
        html_message = MIMEText(body, 'html')
        message.attach(html_message)

        anexos = MIMEBase('application', "octet-stream")
        anexos.set_payload(open(caminho_top_ano, "rb").read())
        encoders.encode_base64(anexos)
        anexos.add_header('Content-Disposition', f'attachment; filename="{unidecode(top_nome_loja_ano)}{data}_TOPAno'
                                                 f'.xlsx"')
        message.attach(anexos)

        anexos = MIMEBase('application', "octet-stream")
        anexos.set_payload(open(caminho_pior_ano, "rb").read())
        encoders.encode_base64(anexos)
        anexos.add_header('Content-Disposition', f'attachment; filename="{unidecode(pior_nome_loja_ano)}{data}_PiorAno'
                                                 f'.xlsx"')
        message.attach(anexos)

        anexos = MIMEBase('application', "octet-stream")
        anexos.set_payload(open(caminho_top_dia, "rb").read())
        encoders.encode_base64(anexos)
        anexos.add_header('Content-Disposition', f'attachment; filename="{unidecode(top_nome_loja_dia)}{data}_TOPDia'
                                                 f'.xlsx"')
        message.attach(anexos)

        anexos = MIMEBase('application', "octet-stream")
        anexos.set_payload(open(caminho_pior_dia, "rb").read())
        encoders.encode_base64(anexos)
        anexos.add_header('Content-Disposition', f'attachment; filename="{unidecode(pior_nome_loja_dia)}{data}_PiorDia'
                                                 f'.xlsx"')
        message.attach(anexos)

        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as server:
            server.login(sender_email, sender_password)
            server.sendmail(sender_email, recipient_email, message.as_string())

        print('E-mail enviado!')
