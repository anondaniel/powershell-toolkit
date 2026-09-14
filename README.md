# powershell-toolkit

Ferramentas de PowerShell para administração de Microsoft 365, Entra ID e
Exchange Online — governança de licenciamento, auditoria de identidade e
offboarding controlado.

Os scripts aqui **vieram de uso em produção**, num tenant de ~1.900 caixas, e
foram parametrizados e sanitizados para funcionar em qualquer ambiente. Não são
exemplos didáticos: carregam as correções que só aparecem depois que algo dá
errado com gente de verdade do outro lado.

---

## O que tem aqui

| Script | O que faz |
|---|---|
| [`Microsoft365/Auditar-Caixas.ps1`](Microsoft365/Auditar-Caixas.ps1) | Cruza a base de RH com Entra ID e Exchange Online e classifica cada caixa por risco: pessoa desligada ainda licenciada, caixa compartilhada com licença desnecessária, conta sem uso. Valida atividade de sign-in antes de sinalizar. Gera relatório HTML. |
| [`Microsoft365/Offboarding-Desligados.ps1`](Microsoft365/Offboarding-Desligados.ps1) | Executa o offboarding em 4 etapas, em ordem não-inversível: converter para compartilhada → renomear → remover licença (grupo **e** direta) → desabilitar. Dry-run por padrão. |
| [`Microsoft365/Scan-Conflitos.ps1`](Microsoft365/Scan-Conflitos.ps1) | Varre caixas procurando lixo de sincronização (pastas de conflito) e calcula o espaço recuperável por caixa. Somente leitura. |

---

## Princípios de desenho

Estes três pontos explicam a maior parte das decisões do código, e todos nasceram
de erro cometido:

**1. Dry-run é o padrão, não uma opção.**
Nenhum script que altera algo faz alteração sem o switch `-Executar`. Rodar sem
ele mostra exatamente o que aconteceria. Não existe "modo automático" implícito.

**2. Nada de tratar como desligado só porque o RH diz.**
Status de RH não é evidência de inatividade: recontratação e mudança de regime
mantêm o registro antigo de demissão. O script valida atividade de sign-in antes
de classificar. Sem essa validação, a primeira versão produzia **141 alertas de
risco alto onde havia 35 reais** — 75% de falso positivo.

**3. A ordem das etapas do offboarding não pode ser invertida.**
Remover a licença de uma caixa de usuário **inicia o prazo de retenção: a caixa
é excluída em ~30 dias**. Converter para compartilhada antes preserva o
histórico e dispensa licença até 50 GB. O script força essa ordem.

Outros cuidados que valem menção:

- **Lista de contas protegidas** no config, que nunca podem ser sinalizadas nem
  tratadas. É a última linha de defesa contra um falso positivo virar incidente.
- **Licença por grupo *e* direta** são tratadas sempre juntas. Remover só do
  grupo deixa a atribuição direta ativa, o custo continua — e o relatório acusa
  sucesso.
- **Contas sincronizadas × cloud-only** seguem caminhos diferentes. Renomear no
  Entra uma conta sincronizada é inútil (o próximo ciclo sobrescreve);
  desabilitar no AD uma conta cloud-only não bloqueia nada.
- **Execução por etapas** (`-Etapas`) permite capturar a economia sem cortar
  acesso, e cortar o acesso depois, numa segunda passada — porque as duas coisas
  têm urgências e aprovadores diferentes.

---

## Pré-requisitos

```powershell
Install-Module Microsoft.Graph -Scope CurrentUser
Install-Module ExchangeOnlineManagement -Scope CurrentUser
Install-Module ImportExcel -Scope CurrentUser   # leitura de planilhas de RH
```

PowerShell 7 é recomendado. **Evite o Windows PowerShell 5.1** para estes
scripts: nele as respostas do Graph vêm como `hashtable` e
`Select-Object -ExpandProperty` falha **em silêncio** — o script conclui
"já sem licença" sem ter removido nada.

### Permissões

Leitura (auditoria): `User.Read.All`, `Directory.Read.All`,
`AuditLog.Read.All` (necessária para `signInActivity`), `Organization.Read.All`.

Escrita (offboarding): `User.ReadWrite.All`, `GroupMember.ReadWrite.All` — mais
uma sessão do Exchange Online e permissão de escrita no AD, quando houver
ambiente híbrido.

Use **apps separados** para leitura e escrita. O app de auditoria não deve ter
permissão de alterar nada.

---

## Configuração

```powershell
Copy-Item config.example.json Microsoft365/config.json
# edite Microsoft365/config.json
```

O `config.json` real está no `.gitignore`. Cada campo está comentado no
[`config.example.json`](config.example.json).

---

## Uso

### Auditoria (somente leitura)

```powershell
# Interativo
.\Auditar-Caixas.ps1 -AdminUPN admin@contoso.com -CSVPath .\colaboradores.csv

# Não-interativo (app-only, lê credenciais do config.json)
.\Auditar-Caixas.ps1
```

### Offboarding

Entrada: um CSV com a coluna `UserPrincipalName`. Veja
[`exemplos/contas-exemplo.csv`](exemplos/contas-exemplo.csv).

```powershell
# 1. Sempre comece em dry-run
.\Offboarding-Desligados.ps1 -ArquivoContas .\sem-uso.csv -Lote sem-uso

# 2. Lote de baixo risco: contas sem uso recente
.\Offboarding-Desligados.ps1 -ArquivoContas .\sem-uso.csv -Lote sem-uso -Executar

# 3. Contas COM uso recente: captura a economia sem cortar acesso
.\Offboarding-Desligados.ps1 -ArquivoContas .\com-uso.csv -Lote com-uso `
    -Etapas Converter,Renomear,Licenca -Executar

# 4. Só depois da confirmação do gestor, corta o acesso
.\Offboarding-Desligados.ps1 -ArquivoContas .\com-uso.csv -Lote com-uso `
    -Etapas Desabilitar -ConfirmarGestor -Executar
```

Separar os lotes por **risco**, e não por conveniência, é o ponto: uma parte das
contas formalmente desligadas ainda está em uso no dia em que você roda o script.

---

## Armadilhas conhecidas

Documentadas porque custaram tempo e nenhuma estava escrita em lugar nenhum:

- **`signInActivity` só vem na listagem de usuários**, com `$select` explícito.
  Um `GET` por UPN não retorna o campo — quem consulta conta a conta conclui que
  o dado não existe.
- **`assignLicense` retorna `BadRequest`** se a lista de remoção incluir um SKU
  que já não está atribuído. É preciso ler o estado antes de cada escrita, e
  remover apenas o que consta em `assignedLicenses`.
- **Add-on não é conflito de licença.** Uma licença complementar atribuída sobre
  um SKU base é o uso correto, não duplicação. Só o acúmulo de dois SKUs *base*
  é conflito.
- **A propagação não é instantânea.** Bloqueio no AD leva cerca de 30 minutos
  para chegar ao Entra por sincronização. Verificar antes disso gera falso
  negativo e retrabalho.
- **Remover membros de grupo exige três cmdlets diferentes** — lista de
  distribuição clássica, grupo do Microsoft 365 e grupo de segurança são coisas
  distintas, e o comando mais óbvio não enxerga os do segundo tipo. A verificação
  confiável é por `memberOf` no Graph, que enxerga os três.

---

## Aviso

Estes scripts alteram contas, licenças e caixas de correio. Rode sempre em
dry-run primeiro, em um lote pequeno, e confira o resultado em uma fonte
independente daquela que o script usou para escrever.

Licença [MIT](LICENSE) — sem garantia de qualquer espécie.
