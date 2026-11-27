using System.Drawing;
using System.Windows.Forms;

namespace OrcamentoNanoCAD.UI
{
    internal sealed class TelaIntercambio : Form
    {
        public string Comando { get; private set; }

        public TelaIntercambio()
        {
            Text = "OrçaCore — Intercâmbio de desenhos";
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            StartPosition = FormStartPosition.CenterParent;
            AutoSize = true;
            AutoSizeMode = AutoSizeMode.GrowAndShrink;
            Padding = new Padding(16);
            Font = Visual.Normal;

            var menu = new TableLayoutPanel { AutoSize = true, ColumnCount = 2 };
            menu.Controls.Add(Visual.Rotulo("Intercâmbio de desenhos", Visual.Titulo), 0, 0);
            Adicionar(menu, 1, "Exportar para JSON", "ORC_EXPORTAR_JSON");
            Adicionar(menu, 2, "Acrescentar ao desenho", "ORC_IMPORTAR_JSON");
            Adicionar(menu, 3, "Gerar DXF separado", "ORC_GERAR_DXF");
            Adicionar(menu, 4, "Abrir DXF", "ORC_ABRIR_DXF");
            Adicionar(menu, 5, "Atualizar geometria", "ORC_ATUALIZAR_DESENHO");
            Adicionar(menu, 6, "Substituir Model Space", "ORC_SUBSTITUIR_DESENHO");
            Adicionar(menu, 7, "Verificar conversor local", "ORC_CONVERSOR_DXF");
            Adicionar(menu, 8, "Exibir ajuda", "ORC_AJUDA_DESENHO");
            Adicionar(menu, 9, "Sincronizar alterações (experimental)", "ORC_SINCRONIZAR_JSON");
            var fechar = Visual.Botao("Fechar", (o, e) => Close());
            fechar.DialogResult = DialogResult.Cancel;
            CancelButton = fechar;
            menu.Controls.Add(fechar, 0, 10);
            Controls.Add(menu);
            Visual.Aplicar(this);
        }

        private void Adicionar(TableLayoutPanel menu, int linha, string rotulo, string comando)
        {
            var botao = Visual.Botao(rotulo, (o, e) =>
            {
                Comando = comando;
                DialogResult = DialogResult.OK;
                Close();
            });
            botao.Anchor = AnchorStyles.Left | AnchorStyles.Right;
            var nome = Visual.Rotulo(comando, Visual.Mono, Visual.Suave);
            nome.Anchor = AnchorStyles.Left;
            nome.Margin = new Padding(12, 3, 3, 3);
            menu.Controls.Add(botao, 0, linha);
            menu.Controls.Add(nome, 1, linha);
        }
    }
}
