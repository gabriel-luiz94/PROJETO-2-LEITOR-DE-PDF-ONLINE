"""tests/oraculo/gerar_dxf_teste.py — monta um DXF de teste parecido com os projetos de Pernambuco (o desenho dos mockups):
rede REDE_MT/REDE_BT, blocos POSTE com atributo, anotações em TEXT e MTEXT (com códigos de cor e vários parágrafos), uma anotação
LONGE do poste ligada por LEADER, um cabo anotado por MLEADER, um rótulo de cabo com o comprimento na linha de baixo e textos que as
regras devem ignorar (ruas, cotas, cor fora do padrão).
Uso: python gerar_dxf_teste.py <saida.dxf>
"""
import sys
import ezdxf
from ezdxf.enums import TextEntityAlignment  # noqa: F401  (garante a versão nova do ezdxf)
from ezdxf.math import Vec2

H = 8.0
doc = ezdxf.new("R2018")
for nome, cor in (("REDE_MT", 1), ("REDE_BT", 3), ("01_INSTALANDO", 7), ("01_REMOVER", 7), ("01_REAPLICANDO", 7), ("TXT_RUAS", 8),
                  ("POSTES", 7), ("GUIAS", 2), ("COTAS", 4), ("TXT_FF0000", 7)):
    doc.layers.add(nome, color=cor)
blk = doc.blocks.new("POSTE")
blk.add_circle((0, 0), 4)
blk.add_attdef("NUM", (5, -5), dxfattribs={"height": 2})
blk.add_text("x", dxfattribs={"height": 1, "insert": (0, 0)})       # texto DENTRO da definição (a aplicação também lê)
msp = doc.modelspace()


def p(x, y):
    return (x, -y)


postes = [(105, 305, "9200D, N1"), (325, 275, None), (555, 335, "11300D, N3"), (785, 285, "9400D, N1"),
          (325, 525, "10300R, N1"), (565, 575, "9300F, N3"), (805, 535, None), (565, 745, "7150FV, B2")]
for i, (x, y, anot) in enumerate(postes):
    ins = msp.add_blockref("POSTE", p(x, y), dxfattribs={"layer": "POSTES"})
    ins.add_attrib("NUM", f"{i + 1}", p(x + 5, y + 5), dxfattribs={"height": 2, "layer": "POSTES"})
    if anot:
        msp.add_text(anot, dxfattribs={"layer": "01_INSTALANDO", "height": H, "insert": p(x - 22, y - 12)})
# P2: MTEXT com 3 parágrafos (o 2º com cor própria entre chaves), perto do poste
msp.add_mtext("9400D\\P{\\C7;N1}\\PN4", dxfattribs={"layer": "01_INSTALANDO", "char_height": H, "insert": p(325 - 22, 275 - 20)})
# P7: anotação LONGE (60 unidades), ligada ao poste por um LEADER na layer GUIAS
msp.add_text("9200D, N1", dxfattribs={"layer": "01_INSTALANDO", "height": H, "insert": p(805 + 40, 535 + 60)})
msp.add_leader([p(805 + 38, 535 + 58), p(805 + 3, 535 + 3)], dxfattribs={"layer": "GUIAS"})


def rede(a, b, tipo, anot, layer_txt, como="text"):
    x1, y1 = postes[a][:2]
    x2, y2 = postes[b][:2]
    msp.add_line(p(x1, y1), p(x2, y2), dxfattribs={"layer": "REDE_MT" if tipo == "MT" else "REDE_BT"})
    mx, my = (x1 + x2) / 2, (y1 + y2) / 2
    vertical = abs(x2 - x1) < abs(y2 - y1)
    tx, ty = (mx - 15, my) if vertical else (mx, my - 12)
    if como == "text":
        msp.add_text(anot, dxfattribs={"layer": layer_txt, "height": H, "insert": p(tx, ty)})
    elif como == "duas_linhas":
        rot, comp = anot.rsplit(" ", 2)[0], " ".join(anot.rsplit(" ", 2)[1:])
        msp.add_text(rot, dxfattribs={"layer": layer_txt, "height": H, "insert": p(tx, ty)})
        msp.add_text(comp, dxfattribs={"layer": layer_txt, "height": H, "insert": p(tx, ty + 10)})
    elif como == "mleader":
        ml = msp.add_multileader_mtext("Standard", dxfattribs={"layer": layer_txt})
        ml.set_content(anot, char_height=H)
        ml.add_leader_line(ezdxf.render.mleader.ConnectionSide.left, [Vec2(p(mx, my))])
        ml.build(insert=Vec2(p(mx + 30, my - 60)))


rede(0, 1, "MT", "3xA35P(7,94)-13,8kV 45 m", "01_INSTALANDO")
rede(1, 2, "MT", "3xA35P(7,94)-13,8kV 45 m", "01_INSTALANDO", como="mleader")
rede(2, 3, "MT", "3xS04-7,97kV 38 m", "01_INSTALANDO", como="duas_linhas")
rede(1, 4, "MT", "3xA35P(7,94)-13,8kV 52 m", "01_REMOVER")
rede(4, 5, "BT", "3xC16(16) 30 m", "01_INSTALANDO")
rede(5, 6, "BT", "3xC16(16) 28 m", "01_INSTALANDO")
rede(5, 7, "BT", "RAMA25 20 m", "01_REAPLICANDO")
# textos que as regras ignoram
msp.add_text("Rua das Flores", dxfattribs={"layer": "TXT_RUAS", "height": H, "insert": p(450, 420)})
msp.add_text("12.50", dxfattribs={"layer": "COTAS", "height": H, "insert": p(200, 200)})
msp.add_mtext("{\\C1;OBS: conferir}", dxfattribs={"layer": "0", "char_height": H, "insert": p(700, 700)})
t = msp.add_text("3xC16(16) 99 m", dxfattribs={"layer": "0", "height": H, "insert": p(900, 900)})
t.rgb = (0, 0, 255)                                                     # azul: regra de motor → operação 0
doc.saveas(sys.argv[1])
