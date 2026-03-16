from datetime import datetime
import csv
import os

class Hob:
    def __init__(self, caminho_csv_temporario, escolha):
        now = datetime.now()
        ano = now.year
        mes = now.month
        dia = now.day
        hora = now.hour
        minuto = now.minute
        segundo = now.second

        if mes == 1:
            mes_ref = 12
            ano_ref = ano - 1
        else:
            mes_ref = mes - 1
            ano_ref = ano

        downloads_path = os.path.join(os.path.expanduser("~"), 'Downloads')
        exit_path = os.path.join(downloads_path, 'MacroFPATMOVFIN_V3_%d%02d%02d%02d%02d%02d.mac' % (ano, mes, dia, hora, minuto, segundo))

        f = open(exit_path, 'a') 
        cabecalho = '''<HAScript name="FPATMOVFINV3" description="" timeout="60000" pausetime="300" promptall="true" blockinput="false" author="marcus.gabaldo" creationdate="%02d/%02d/%d %02d:%02d:%02d" supressclearevents="false" usevars="false" ignorepauseforenhancedtn="true" delayifnotenhancedtn="0" ignorepausetimeforenhancedtn="true">''' % (dia, mes, ano, hora, minuto, segundo)
        f.write(cabecalho) 

        input_file = open(caminho_csv_temporario, 'r', newline='')

        data = csv.reader(input_file)
        i = 1

        corpo = '''
            <screen name="Tela%d" entryscreen="true" exitscreen="false" transient="false">
                <description >
                    <oia status="NOTINHIBITED" optional="false" invertmatch="false" />
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
                    <numfields number="82" optional="false" invertmatch="false" />
                    <numinputfields number="6" optional="false" invertmatch="false" />
                </description>
                <actions>
                    <input value="%s%s%s%sX[enter]" row="0" col="0" movecursor="true" xlatehostkeys="true" encrypted="false" />
                </actions>
                <nextscreens timeout="0" >
                    <nextscreen name="Tela%d" />
                </nextscreens>
            </screen>

            <screen name="Tela%d" entryscreen="false" exitscreen="false" transient="false">
                <description >
                    <oia status="NOTINHIBITED" optional="false" invertmatch="false" />
                    <numfields number="390" optional="false" invertmatch="false" />
                    <numinputfields number="101" optional="false" invertmatch="false" />
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
                    <numfields number="391" optional="false" invertmatch="false" />
                    <numinputfields number="101" optional="false" invertmatch="false" />
                </description>
                <actions>
                    <mouseclick row="16" col="26" />
                    <input value="%s[tab]%s[enter]" row="0" col="0" movecursor="true" xlatehostkeys="true" encrypted="false" />
                </actions>
                <nextscreens timeout="0" >
                    <nextscreen name="Tela%d" />
                </nextscreens>
            </screen>

            <screen name="Tela%d" entryscreen="false" exitscreen="false" transient="false">
                <description >
                    <oia status="NOTINHIBITED" optional="false" invertmatch="false" />
                    <numfields number="86" optional="false" invertmatch="false" />
                    <numinputfields number="1" optional="false" invertmatch="false" />
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
                    <numfields number="93" optional="false" invertmatch="false" />
                    <numinputfields number="1" optional="false" invertmatch="false" />
                </description>
                <actions>
                    <input value="c[enter]" row="0" col="0" movecursor="true" xlatehostkeys="true" encrypted="false" />
                </actions>
                <nextscreens timeout="0" >
                    <nextscreen name="Tela%d" />
                </nextscreens>
            </screen>

            <screen name="Tela%d" entryscreen="false" exitscreen="false" transient="false">
                <description >
                    <oia status="NOTINHIBITED" optional="false" invertmatch="false" />
                    <numfields number="88" optional="false" invertmatch="false" />
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
                <numfields number="82" optional="false" invertmatch="false" />
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

        if escolha == 'VT':
            valor_tela3 = f"82695"
        else:
            valor_tela3 = f"83172"

        for line in data: 
            [matricula, rendesc, rubrica, sequencia, iae, ref, valor, doclegal, justificativa] = line

            tela1 = i
            tela2 = tela1 + 1
            tela3 = tela2 + 1
            tela4 = tela3 + 1
            tela5 = tela4 + 1
            tela6 = tela5 + 1
            tela7 = tela6 + 1
            tela8 = tela7 + 1
            tela9 = tela8 + 1

            f.write('\n')
            f.write(corpo % (
                tela1, matricula, tela2, 
                tela2, rendesc, rubrica, sequencia, iae, tela3, 
                tela3, f"{valor_tela3}01{mes_ref:02d}{ano_ref}01{mes_ref:02d}{ano_ref}{valor}", tela4, 
                tela4, doclegal, justificativa, tela5, 
                tela5, ref, tela6, 
                tela6, tela7, 
                tela7, tela8,
                tela8, tela9
            ))

            i = i + 8

        f.write('\n')
        f.write(rodape % (i))
        f.close() 
        input_file.close()