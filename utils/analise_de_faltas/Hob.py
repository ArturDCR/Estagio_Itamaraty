# Macro para SERPRO by Marcus Gabaldo 2014
from datetime import datetime
import csv
import os

class Hob:
    def __init__(self):
        now = datetime.now()
        ano = (now.year)
        mes = (now.month)
        dia = (now.day)
        hora = (now.hour)
        minuto = (now.minute)
        segundo = (now.second)

        downloads_path = os.path.join(os.path.expanduser("~"), 'Downloads')
        exit_path = os.path.join(downloads_path, 'MacroFPATMOVFIN_V3_%d%02d%02d%02d%02d%02d.mac' % (ano, mes, dia, hora, minuto, segundo))

        f = open(exit_path, 'a') # abre o arquivo para escrever, o 'a' serve para adicionar

        cabecalho = '''<HAScript name="FPATMOVFINV3" description="" timeout="60000" pausetime="300" promptall="true" blockinput="false" author="marcus.gabaldo" creationdate="04/09/2014 11:51:31" supressclearevents="false" usevars="false" ignorepauseforenhancedtn="true" delayifnotenhancedtn="0" ignorepausetimeforenhancedtn="true">'''
        f.write(cabecalho) # escreve o cabecalho no arquivo

        input_file = open('utils/analise_de_faltas/dados/dadosFPATMOVFIN_V3_REF.csv','r',newline='')

        data = csv.reader(input_file)

        i = 1

        for line in data: # verifica cada linha do arquivo csv
            [matricula, rendesc, rubrica, sequencia, iae, ref, valor, doclegal, justificativa] = line

            corpo = '''
            <screen name="Tela%d" entryscreen="false" exitscreen="false" transient="false">
                <description >
                    <oia status="NOTINHIBITED" optional="false" invertmatch="false" />
                    <numfields number="70" optional="false" invertmatch="false" />
                    <numinputfields number="5" optional="false" invertmatch="false" />
                </description>
                <actions>
                    <input value="%s[enter]" row="0" col="0" movecursor="true" xlatehostkeys="true" encrypted="false" />
                </actions>
                <nextscreens timeout="0" >
                    <nextscreen name="Tela%d" />
                </nextscreens>
            </screen>

            <screen name="Tela%d" entryscreen="false" exitscreen="false" transient="false">
                <description >
                    <oia status="NOTINHIBITED" optional="false" invertmatch="false" />
                    <numfields number="84" optional="false" invertmatch="false" />
                    <numinputfields number="6" optional="false" invertmatch="false" />
                </description>
                <actions>
                    <input value="%s%s%s%s[tab]X[enter]" row="0" col="0" movecursor="true" xlatehostkeys="true" encrypted="false" />
                </actions>
                <nextscreens timeout="0" >
                    <nextscreen name="Tela%d" />
                </nextscreens>
            </screen>

            <screen name="Tela%d" entryscreen="false" exitscreen="false" transient="false">
                <description >
                    <oia status="NOTINHIBITED" optional="false" invertmatch="false" />
                    <numfields number="134" optional="false" invertmatch="false" />
                    <numinputfields number="17" optional="false" invertmatch="false" />
                </description>
                <actions>
                    <input value="%s%s44[enter]" row="0" col="0" movecursor="true" xlatehostkeys="true" encrypted="false" />
                </actions>
                <nextscreens timeout="0" >
                    <nextscreen name="Tela%d" />
                </nextscreens>
            </screen>

            <screen name="Tela%d" entryscreen="false" exitscreen="false" transient="false">
                <description >
                    <oia status="NOTINHIBITED" optional="false" invertmatch="false" />
                    <numfields number="134" optional="false" invertmatch="false" />
                    <numinputfields number="7" optional="false" invertmatch="false" />
                </description>
                <actions>
                    <input value="%s[tab]%s[enter]" row="0" col="0" movecursor="true" xlatehostkeys="true" encrypted="false" />
                </actions>
                <nextscreens timeout="0" >
                    <nextscreen name="Tela%d" />
                </nextscreens>
            </screen>

            <screen name="Tela%d" entryscreen="false" exitscreen="false" transient="false">
                <description >
                    <oia status="NOTINHIBITED" optional="false" invertmatch="false" />
                    <numfields number="52" optional="false" invertmatch="false" />
                    <numinputfields number="1" optional="false" invertmatch="false" />
                </description>
                <actions>
                    <input value="C[enter]" row="0" col="0" movecursor="true" xlatehostkeys="true" encrypted="false" />
                </actions>
                <nextscreens timeout="0" >
                    <nextscreen name="Tela%d" />
                </nextscreens>
            </screen>

            <screen name="Tela%d" entryscreen="false" exitscreen="false" transient="false">
                <description >
                    <oia status="NOTINHIBITED" optional="false" invertmatch="false" />
                    <numfields number="53" optional="false" invertmatch="false" />
                    <numinputfields number="0" optional="false" invertmatch="false" />
                </description>
                <actions>
                    <input value="[enter]" row="0" col="0" movecursor="true" xlatehostkeys="true" encrypted="false" />
                </actions>
                <nextscreens timeout="0" >
                    <nextscreen name="Tela%d" />
                </nextscreens>
            </screen>

            <screen name="Tela%d" entryscreen="false" exitscreen="false" transient="false">
                <description >
                    <oia status="NOTINHIBITED" optional="false" invertmatch="false" />
                    <numfields number="84" optional="false" invertmatch="false" />
                    <numinputfields number="6" optional="false" invertmatch="false" />
                </description>
                <actions>
                    <input value="[pf12]" row="0" col="0" movecursor="true" xlatehostkeys="true" encrypted="false" />
                </actions>
                <nextscreens timeout="0" >
                    <nextscreen name="Tela%d" />
                </nextscreens>
            </screen>'''

            rodape = '''
            <screen name="Tela%d" entryscreen="false" exitscreen="true" transient="false">
                <description >
                    <oia status="NOTINHIBITED" optional="false" invertmatch="false" />
                    <numfields number="70" optional="false" invertmatch="false" />
                    <numinputfields number="5" optional="false" invertmatch="false" />
                </description>
                <actions>
                    <input value="[pf3]" row="0" col="0" movecursor="true" xlatehostkeys="true" encrypted="false" />
                </actions>
                <nextscreens timeout="0" >
                </nextscreens>
            </screen>

        </HAScript>'''

            tela1 = i
            tela2 = tela1 + 1
            tela3 = tela2 + 1
            tela4 = tela3 + 1
            tela5 = tela4 + 1
            tela6 = tela5 + 1
            tela7 = tela6 + 1
            tela8 = tela7 + 1

            f.write('\n')
            f.write(corpo % (tela1, matricula, tela2, tela2, rendesc, rubrica, sequencia, iae, tela3, tela3, ref, valor, tela4, tela4, doclegal, justificativa, tela5, tela5, tela6, tela6, tela7, tela7, tela8))

            i = i + 7

        f.write('\n')
        f.write(rodape % (i))# escreve o texto no arquivo
        f.close() # fecha o arquivo