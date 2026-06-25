// ===== Code.gs =====
const NOME_SISTEMA = 'Fila Cirúrgica ORL';

function doGet() {
  inicializarPlanilha();
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle(NOME_SISTEMA)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(nomeArquivo) {
  return HtmlService.createHtmlOutputFromFile(nomeArquivo).getContent();
}

function obterSpreadsheet() {
  const props = PropertiesService.getScriptProperties();
  const spreadsheetId = props.getProperty('SPREADSHEET_ID');
  if (spreadsheetId) {
    return SpreadsheetApp.openById(spreadsheetId);
  }

  const ativa = SpreadsheetApp.getActiveSpreadsheet();
  if (ativa) {
    props.setProperty('SPREADSHEET_ID', ativa.getId());
    return ativa;
  }

  const criada = SpreadsheetApp.create('Fila Cirúrgica ORL - Dados');
  props.setProperty('SPREADSHEET_ID', criada.getId());
  return criada;
}

function inicializarPlanilha() {
  const ss = obterSpreadsheet();
  criarAbaSeNecessario_(ss, 'Pacientes', CABECALHOS.Pacientes);
  criarAbaSeNecessario_(ss, 'Hospitais', CABECALHOS.Hospitais);
  criarAbaSeNecessario_(ss, 'Usuários', CABECALHOS.Usuarios);
  criarAbaSeNecessario_(ss, 'Histórico', CABECALHOS.Historico);
  criarAbaSeNecessario_(ss, 'ModelosWhatsApp', CABECALHOS.ModelosWhatsApp);
  criarAbaSeNecessario_(ss, 'PlanosDeSaúde', CABECALHOS.PlanosDeSaude);
  criarAbaSeNecessario_(ss, 'Configurações', CABECALHOS.Configuracoes);

  semearHospitais_();
  semearUsuarios_();
  semearModelosWhatsApp_();
  semearConfiguracoes_();
  aplicarValidacoesPacientes_();

  return {
    nome: ss.getName(),
    id: ss.getId(),
    url: ss.getUrl()
  };
}

function obterDadosIniciais(sessao) {
  const usuario = validarSessao_(sessao);
  const dados = {
    usuario: usuario,
    configuracoes: obterConfiguracoes(),
    pacientes: obterPacientes(usuario),
    hospitais: obterHospitais(usuario),
    planosSaude: obterPlanosSaude(usuario),
    modelosWhatsApp: obterModelosWhatsApp(usuario),
    resumo: calcularResumo(obterPacientes(usuario)),
    statusDisponiveis: STATUS_DISPONIVEIS,
    proximasAcoes: PROXIMAS_ACOES
  };

  if (usuario.Perfil === 'Administrador') {
    dados.usuarios = obterUsuarios(usuario);
  }

  return dados;
}

function criarAbaSeNecessario_(ss, nome, cabecalhos) {
  let sheet = ss.getSheetByName(nome);
  if (!sheet) {
    sheet = ss.insertSheet(nome);
  }

  const range = sheet.getRange(1, 1, 1, cabecalhos.length);
  const atuais = range.getValues()[0];
  const precisaCabecalho = atuais.every(valor => valor === '');

  if (precisaCabecalho) {
    range.setValues([cabecalhos]);
    range.setFontWeight('bold');
    sheet.setFrozenRows(1);
  }

  return sheet;
}

function obterAba_(nome) {
  return obterSpreadsheet().getSheetByName(nome);
}

function lerRegistros_(nomeAba) {
  const sheet = obterAba_(nomeAba);
  if (!sheet) return [];
  const valores = sheet.getDataRange().getValues();
  if (valores.length <= 1) return [];

  const cabecalhos = valores[0];
  return valores.slice(1)
    .filter(linha => linha.some(valor => valor !== ''))
    .map((linha, index) => {
      const registro = { _linha: index + 2 };
      cabecalhos.forEach((cabecalho, posicao) => {
        registro[cabecalho] = normalizarValorSaida_(linha[posicao]);
      });
      return registro;
    });
}

function salvarLinha_(nomeAba, registro) {
  const sheet = obterAba_(nomeAba);
  const cabecalhos = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const linha = cabecalhos.map(cabecalho => registro[cabecalho] === undefined ? '' : registro[cabecalho]);
  sheet.appendRow(linha);
  return sheet.getLastRow();
}

function atualizarLinha_(nomeAba, numeroLinha, registro) {
  const sheet = obterAba_(nomeAba);
  const cabecalhos = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const linha = cabecalhos.map(cabecalho => registro[cabecalho] === undefined ? '' : registro[cabecalho]);
  sheet.getRange(numeroLinha, 1, 1, linha.length).setValues([linha]);
}

function agora_() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
}

function hoje_() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function gerarId_() {
  return Utilities.getUuid();
}

function normalizarValorSaida_(valor) {
  if (Object.prototype.toString.call(valor) === '[object Date]' && !isNaN(valor)) {
    return Utilities.formatDate(valor, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return valor === null || valor === undefined ? '' : valor;
}

function diasEntre_(dataTexto) {
  if (!dataTexto) return null;
  const data = new Date(dataTexto);
  if (isNaN(data.getTime())) return null;
  const hoje = new Date();
  hoje.setHours(0, 0, 0, 0);
  data.setHours(0, 0, 0, 0);
  return Math.floor((hoje.getTime() - data.getTime()) / 86400000);
}

function sanitizarTexto_(valor) {
  return String(valor || '').trim();
}

function validarSessao_(sessao) {
  if (!sessao || !sessao.usuario) {
    throw new Error('Sessão inválida. Faça login novamente.');
  }

  const usuario = obterUsuarioPorNome(sessao.usuario);
  if (!usuario || usuario.Ativo !== 'Sim') {
    throw new Error('Usuário sem permissão ou inativo.');
  }
  return usuario;
}

function criarResposta_(sucesso, dados, mensagem) {
  return { sucesso: sucesso, dados: dados || null, mensagem: mensagem || '' };
}

// ===== Configuracoes.gs =====
const CABECALHOS = {
  Pacientes: ['ID', 'Nome', 'DataNascimento', 'Telefone', 'Email', 'PlanoSaúde', 'TipoPagamento', 'Status', 'DataCirurgia', 'HoraCirurgia', 'Hospital', 'Observações', 'PróximaAção', 'DataPróximoContato', 'CriadoEm', 'CriadoPor', 'AtualizadoEm', 'AtualizadoPor', 'Arquivado'],
  Hospitais: ['ID', 'Nome', 'CriadoEm', 'CriadoPor', 'Ativo'],
  Usuarios: ['Usuário', 'Perfil', 'Ativo'],
  Historico: ['ID', 'IDPaciente', 'NomePaciente', 'Usuario', 'PerfilUsuario', 'TipoAção', 'CampoAlterado', 'ValorAnterior', 'ValorNovo', 'Observação', 'CriadoEm'],
  ModelosWhatsApp: ['ID', 'Título', 'Mensagem', 'Ativo', 'CriadoEm'],
  PlanosDeSaude: ['ID', 'Nome', 'Ativo', 'CriadoEm', 'CriadoPor'],
  Configuracoes: ['Chave', 'Valor']
};

const STATUS_DISPONIVEIS = [
  'Aguardando tomografia',
  'Aguardando exames/documentos',
  'Solicitado ao plano de saúde',
  'Em análise pelo plano',
  'Autorizado',
  'Aguardando agendamento',
  'Agendado',
  'Realizado',
  'Pendente / travado',
  'Cancelado'
];

const PROXIMAS_ACOES = [
  'Cobrar tomografia',
  'Cobrar exames pré-operatórios',
  'Aguardar envio de documentos',
  'Enviar solicitação ao plano',
  'Cobrar resposta do plano',
  'Confirmar autorização',
  'Agendar cirurgia',
  'Confirmar hospital',
  'Confirmar pagamento particular',
  'Enviar orientações pré-operatórias',
  'Confirmar presença',
  'Marcar pós-operatório',
  'Sem ação pendente'
];

const CONFIGURACOES_INICIAIS = [
  ['NomeSistema', 'Fila Cirúrgica ORL'],
  ['CodigoAcesso', 'septo123'],
  ['DiasSemAtualizaçãoAlerta', '7'],
  ['DiasSolicitadoPlanoAlerta', '14'],
  ['DiasAutorizadoSemAgendamentoAlerta', '7']
];

function obterConfiguracoes() {
  const registros = lerRegistros_('Configurações');
  return registros.reduce((acc, item) => {
    acc[item.Chave] = item.Valor;
    return acc;
  }, {});
}

function obterCodigoAcesso() {
  return obterConfiguracoes().CodigoAcesso || 'septo123';
}

function alterarCodigoAcesso(novoCodigo, sessao) {
  const usuario = validarSessao_(sessao);
  verificarPermissao(usuario, 'alterarCodigoAcesso');

  const codigo = sanitizarTexto_(novoCodigo);
  if (!codigo) throw new Error('Informe um novo código de acesso.');

  const sheet = obterAba_('Configurações');
  const registros = lerRegistros_('Configurações');
  const existente = registros.find(item => item.Chave === 'CodigoAcesso');
  const anterior = existente ? existente.Valor : '';

  if (existente) {
    sheet.getRange(existente._linha, 2).setValue(codigo);
  } else {
    sheet.appendRow(['CodigoAcesso', codigo]);
  }

  registrarHistorico({
    usuario: usuario,
    tipoAcao: 'Alteração de configuração',
    campoAlterado: 'CodigoAcesso',
    valorAnterior: anterior ? 'Código anterior' : '',
    valorNovo: 'Novo código definido',
    observacao: 'Código de acesso alterado.'
  });

  return criarResposta_(true, null, 'Código de acesso alterado.');
}

function semearConfiguracoes_() {
  const registros = lerRegistros_('Configurações');
  const chaves = registros.map(item => item.Chave);
  CONFIGURACOES_INICIAIS.forEach(item => {
    if (!chaves.includes(item[0])) {
      salvarLinha_('Configurações', { Chave: item[0], Valor: item[1] });
    }
  });
}

// ===== Usuarios.gs =====
const USUARIOS_INICIAIS = [
  ['Felipe', 'Administrador', 'Sim'],
  ['João', 'Administrador', 'Sim'],
  ['Raquel', 'Secretária', 'Sim'],
  ['Lizandra', 'Secretária', 'Sim'],
  ['Livia', 'Secretária', 'Sim']
];

const PERMISSOES = {
  Administrador: ['cadastrarPaciente', 'editarPaciente', 'arquivarPaciente', 'alterarStatus', 'cadastrarHospital', 'cadastrarPlano', 'gerenciarUsuarios', 'verHistoricoCompleto', 'verHistoricoPaciente', 'editarModelosWhatsApp', 'usarWhatsApp', 'alterarCodigoAcesso'],
  Secretária: ['cadastrarPaciente', 'editarPaciente', 'alterarStatus', 'cadastrarHospital', 'cadastrarPlano', 'usarWhatsApp', 'verHistoricoPaciente']
};

function validarLogin(usuarioNome, codigo) {
  inicializarPlanilha();
  const nome = sanitizarTexto_(usuarioNome);
  const usuario = obterUsuarioPorNome(nome);

  if (!usuario || usuario.Ativo !== 'Sim' || String(codigo || '') !== String(obterCodigoAcesso())) {
    return criarResposta_(false, null, 'Usuário ou código de acesso inválido.');
  }

  registrarHistorico({
    usuario: usuario,
    tipoAcao: 'Login realizado',
    observacao: 'Login realizado no aplicativo.'
  });

  return criarResposta_(true, {
    usuario: usuario['Usuário'],
    perfil: usuario.Perfil,
    loginEm: agora_()
  }, 'Login realizado.');
}

function obterUsuarioPorNome(usuarioNome) {
  const nome = sanitizarTexto_(usuarioNome).toLowerCase();
  if (!nome) return null;
  return lerRegistros_('Usuários').find(usuario => String(usuario['Usuário']).toLowerCase() === nome) || null;
}

function obterUsuarios(sessaoOuUsuario) {
  const usuario = sessaoOuUsuario && sessaoOuUsuario.Perfil ? sessaoOuUsuario : validarSessao_(sessaoOuUsuario);
  verificarPermissao(usuario, 'gerenciarUsuarios');
  return lerRegistros_('Usuários');
}

function salvarUsuario(dados, sessao) {
  const usuarioLogado = validarSessao_(sessao);
  verificarPermissao(usuarioLogado, 'gerenciarUsuarios');

  const nome = sanitizarTexto_(dados['Usuário'] || dados.usuario);
  const perfil = sanitizarTexto_(dados.Perfil || dados.perfil || 'Secretária');
  const ativo = sanitizarTexto_(dados.Ativo || dados.ativo || 'Sim');

  if (!nome) throw new Error('Informe o nome do usuário.');
  if (!['Administrador', 'Secretária'].includes(perfil)) throw new Error('Perfil inválido.');

  const existente = obterUsuarioPorNome(nome);
  if (existente) {
    const anterior = Object.assign({}, existente);
    existente['Usuário'] = nome;
    existente.Perfil = perfil;
    existente.Ativo = ativo === 'Não' ? 'Não' : 'Sim';
    atualizarLinha_('Usuários', existente._linha, existente);
    registrarHistorico({
      usuario: usuarioLogado,
      tipoAcao: 'Edição de usuário',
      campoAlterado: 'Usuário',
      valorAnterior: JSON.stringify(anterior),
      valorNovo: JSON.stringify({ nome: nome, perfil: perfil, ativo: existente.Ativo }),
      observacao: 'Usuário atualizado.'
    });
  } else {
    salvarLinha_('Usuários', { 'Usuário': nome, Perfil: perfil, Ativo: ativo === 'Não' ? 'Não' : 'Sim' });
    registrarHistorico({
      usuario: usuarioLogado,
      tipoAcao: 'Cadastro de usuário',
      valorNovo: nome,
      observacao: 'Usuário cadastrado.'
    });
  }

  return criarResposta_(true, obterUsuarios(usuarioLogado), 'Usuário salvo.');
}

function desativarUsuario(nomeUsuario, sessao) {
  const usuarioLogado = validarSessao_(sessao);
  verificarPermissao(usuarioLogado, 'gerenciarUsuarios');

  const usuario = obterUsuarioPorNome(nomeUsuario);
  if (!usuario) throw new Error('Usuário não encontrado.');
  usuario.Ativo = 'Não';
  atualizarLinha_('Usuários', usuario._linha, usuario);

  registrarHistorico({
    usuario: usuarioLogado,
    tipoAcao: 'Desativação de usuário',
    campoAlterado: 'Ativo',
    valorAnterior: 'Sim',
    valorNovo: 'Não',
    observacao: 'Usuário desativado.'
  });

  return criarResposta_(true, obterUsuarios(usuarioLogado), 'Usuário desativado.');
}

function verificarPermissao(usuario, acao) {
  const perfil = usuario.Perfil || usuario.perfil;
  const permitido = PERMISSOES[perfil] && PERMISSOES[perfil].includes(acao);
  if (!permitido) {
    throw new Error('Você não tem permissão para esta ação.');
  }
  return true;
}

function semearUsuarios_() {
  const existentes = lerRegistros_('Usuários').map(item => String(item['Usuário']).toLowerCase());
  USUARIOS_INICIAIS.forEach(item => {
    if (!existentes.includes(item[0].toLowerCase())) {
      salvarLinha_('Usuários', { 'Usuário': item[0], Perfil: item[1], Ativo: item[2] });
    }
  });
}

// ===== Historico.gs =====
function registrarHistorico(opcoes) {
  const usuario = opcoes.usuario || {};
  salvarLinha_('Histórico', {
    ID: gerarId_(),
    IDPaciente: opcoes.idPaciente || '',
    NomePaciente: opcoes.nomePaciente || '',
    Usuario: usuario['Usuário'] || usuario.usuario || '',
    PerfilUsuario: usuario.Perfil || usuario.perfil || '',
    TipoAção: opcoes.tipoAcao || '',
    CampoAlterado: opcoes.campoAlterado || '',
    ValorAnterior: opcoes.valorAnterior || '',
    ValorNovo: opcoes.valorNovo || '',
    Observação: opcoes.observacao || '',
    CriadoEm: agora_()
  });
}

function obterHistoricoPaciente(idPaciente, sessao) {
  const usuario = validarSessao_(sessao);
  verificarPermissao(usuario, 'verHistoricoPaciente');
  return lerRegistros_('Histórico')
    .filter(item => item.IDPaciente === idPaciente)
    .sort((a, b) => String(b.CriadoEm).localeCompare(String(a.CriadoEm)));
}

function obterHistoricoGeral(sessao) {
  const usuario = validarSessao_(sessao);
  verificarPermissao(usuario, 'verHistoricoCompleto');
  return lerRegistros_('Histórico')
    .sort((a, b) => String(b.CriadoEm).localeCompare(String(a.CriadoEm)))
    .slice(0, 500);
}

function registrarMudancasPaciente_(usuario, anterior, atual) {
  const camposIgnorados = ['_linha'];
  Object.keys(atual).forEach(campo => {
    if (camposIgnorados.includes(campo)) return;
    const antes = anterior ? String(anterior[campo] || '') : '';
    const depois = String(atual[campo] || '');
    if (antes !== depois) {
      registrarHistorico({
        usuario: usuario,
        idPaciente: atual.ID,
        nomePaciente: atual.Nome,
        tipoAcao: campo === 'Status' ? 'Mudança de status' : 'Edição de dados',
        campoAlterado: campo,
        valorAnterior: antes,
        valorNovo: depois,
        observacao: 'Paciente atualizado.'
      });
    }
  });
}

// ===== Pacientes.gs =====
function obterPacientes(sessaoOuUsuario) {
  const usuario = sessaoOuUsuario && sessaoOuUsuario.Perfil ? sessaoOuUsuario : validarSessao_(sessaoOuUsuario);
  const pacientes = lerRegistros_('Pacientes');
  const configuracoes = obterConfiguracoes();

  return pacientes.map(paciente => {
    paciente.PrecisaAtenção = pacientePrecisaAtencao_(paciente, configuracoes);
    paciente.MotivosAtenção = motivosAtencao_(paciente, configuracoes);
    return paciente;
  }).filter(paciente => usuario.Perfil === 'Administrador' || paciente.Arquivado !== 'Sim');
}

function salvarPaciente(dados, sessao) {
  const usuario = validarSessao_(sessao);
  verificarPermissao(usuario, 'cadastrarPaciente');
  const paciente = montarPaciente_(dados, null, usuario);
  validarPaciente_(paciente);
  paciente.ID = gerarId_();
  paciente.CriadoEm = agora_();
  paciente.CriadoPor = usuario['Usuário'];
  paciente.AtualizadoEm = agora_();
  paciente.AtualizadoPor = usuario['Usuário'];
  paciente.Arquivado = 'Não';

  salvarPlanoSeNovo_(paciente.PlanoSaúde, usuario);
  salvarLinha_('Pacientes', paciente);
  registrarHistorico({
    usuario: usuario,
    idPaciente: paciente.ID,
    nomePaciente: paciente.Nome,
    tipoAcao: 'Cadastro de paciente',
    valorNovo: paciente.Nome,
    observacao: 'Paciente cadastrado.'
  });

  return criarResposta_(true, paciente, 'Paciente cadastrado.');
}

function atualizarPaciente(dados, sessao) {
  const usuario = validarSessao_(sessao);
  verificarPermissao(usuario, 'editarPaciente');

  const anterior = encontrarPaciente_(dados.ID);
  if (!anterior) throw new Error('Paciente não encontrado.');

  const paciente = montarPaciente_(dados, anterior, usuario);
  validarPaciente_(paciente);
  paciente.ID = anterior.ID;
  paciente.CriadoEm = anterior.CriadoEm;
  paciente.CriadoPor = anterior.CriadoPor;
  paciente.AtualizadoEm = agora_();
  paciente.AtualizadoPor = usuario['Usuário'];
  paciente.Arquivado = anterior.Arquivado || 'Não';

  salvarPlanoSeNovo_(paciente.PlanoSaúde, usuario);
  atualizarLinha_('Pacientes', anterior._linha, paciente);
  registrarMudancasPaciente_(usuario, anterior, paciente);

  return criarResposta_(true, paciente, 'Paciente atualizado.');
}

function arquivarPaciente(idPaciente, sessao) {
  const usuario = validarSessao_(sessao);
  verificarPermissao(usuario, 'arquivarPaciente');
  const paciente = encontrarPaciente_(idPaciente);
  if (!paciente) throw new Error('Paciente não encontrado.');

  const anterior = paciente.Arquivado || 'Não';
  paciente.Arquivado = 'Sim';
  paciente.AtualizadoEm = agora_();
  paciente.AtualizadoPor = usuario['Usuário'];
  atualizarLinha_('Pacientes', paciente._linha, paciente);

  registrarHistorico({
    usuario: usuario,
    idPaciente: paciente.ID,
    nomePaciente: paciente.Nome,
    tipoAcao: 'Arquivamento',
    campoAlterado: 'Arquivado',
    valorAnterior: anterior,
    valorNovo: 'Sim',
    observacao: 'Paciente arquivado.'
  });

  return criarResposta_(true, paciente, 'Paciente arquivado.');
}

function alterarStatusPaciente(idPaciente, status, dataCirurgia, hospital, horaCirurgia, sessao) {
  const usuario = validarSessao_(sessao);
  verificarPermissao(usuario, 'alterarStatus');
  const paciente = encontrarPaciente_(idPaciente);
  if (!paciente) throw new Error('Paciente não encontrado.');

  const anterior = Object.assign({}, paciente);
  paciente.Status = sanitizarTexto_(status);
  paciente.DataCirurgia = dataCirurgia || paciente.DataCirurgia || '';
  paciente.HoraCirurgia = horaCirurgia || paciente.HoraCirurgia || '';
  paciente.Hospital = hospital || paciente.Hospital || '';
  paciente.AtualizadoEm = agora_();
  paciente.AtualizadoPor = usuario['Usuário'];
  validarPaciente_(paciente);
  atualizarLinha_('Pacientes', paciente._linha, paciente);
  registrarMudancasPaciente_(usuario, anterior, paciente);

  return criarResposta_(true, paciente, 'Status alterado.');
}

function calcularResumo(pacientes) {
  const lista = pacientes || lerRegistros_('Pacientes').filter(item => item.Arquivado !== 'Sim');
  const hoje = hoje_();
  const resumo = {
    total: 0,
    aguardandoTomografia: 0,
    solicitadosPlano: 0,
    autorizados: 0,
    agendados: 0,
    pendentesTravados: 0,
    proximosHoje: 0,
    proximosAtrasados: 0
  };

  lista.forEach(paciente => {
    if (paciente.Arquivado === 'Sim') return;
    resumo.total++;
    if (paciente.Status === 'Aguardando tomografia') resumo.aguardandoTomografia++;
    if (paciente.Status === 'Solicitado ao plano de saúde') resumo.solicitadosPlano++;
    if (paciente.Status === 'Autorizado') resumo.autorizados++;
    if (paciente.Status === 'Agendado') resumo.agendados++;
    if (paciente.Status === 'Pendente / travado') resumo.pendentesTravados++;
    if (paciente.DataPróximoContato === hoje) resumo.proximosHoje++;
    if (paciente.DataPróximoContato && paciente.DataPróximoContato < hoje) resumo.proximosAtrasados++;
  });

  return resumo;
}

function filtrarPacientes(pacientes, filtro) {
  const hoje = hoje_();
  return pacientes.filter(paciente => {
    switch (filtro) {
      case 'Precisa de atenção': return paciente.PrecisaAtenção;
      case 'Próximos contatos hoje': return paciente.DataPróximoContato === hoje;
      case 'Próximos contatos atrasados': return paciente.DataPróximoContato && paciente.DataPróximoContato < hoje;
      case 'Sem próxima ação definida': return !paciente.PróximaAção;
      case 'Particulares': return paciente.TipoPagamento === 'Particular';
      case 'Convênios': return paciente.TipoPagamento === 'Convênio';
      case 'Arquivados': return paciente.Arquivado === 'Sim';
      case 'Pendentes/travados': return paciente.Status === 'Pendente / travado';
      case 'Todos': return paciente.Arquivado !== 'Sim';
      default: return paciente.Status === filtro && paciente.Arquivado !== 'Sim';
    }
  });
}

function encontrarPaciente_(idPaciente) {
  return lerRegistros_('Pacientes').find(item => item.ID === idPaciente) || null;
}

function montarPaciente_(dados, anterior, usuario) {
  const base = anterior || {};
  const tipoPagamento = sanitizarTexto_(dados.TipoPagamento || base.TipoPagamento);
  let plano = sanitizarTexto_(dados.PlanoSaúde || base.PlanoSaúde);
  if (tipoPagamento === 'Particular' && !plano) plano = 'Particular';

  return {
    ID: base.ID || '',
    Nome: sanitizarTexto_(dados.Nome || base.Nome),
    DataNascimento: dados.DataNascimento || base.DataNascimento || '',
    Telefone: sanitizarTexto_(dados.Telefone || base.Telefone),
    Email: sanitizarTexto_(dados.Email || base.Email),
    PlanoSaúde: plano,
    TipoPagamento: tipoPagamento,
    Status: sanitizarTexto_(dados.Status || base.Status),
    DataCirurgia: dados.DataCirurgia || base.DataCirurgia || '',
    HoraCirurgia: dados.HoraCirurgia || base.HoraCirurgia || '',
    Hospital: sanitizarTexto_(dados.Hospital || base.Hospital),
    Observações: sanitizarTexto_(dados.Observações || base.Observações),
    PróximaAção: sanitizarTexto_(dados.PróximaAção || base.PróximaAção),
    DataPróximoContato: dados.DataPróximoContato || base.DataPróximoContato || '',
    CriadoEm: base.CriadoEm || '',
    CriadoPor: base.CriadoPor || '',
    AtualizadoEm: agora_(),
    AtualizadoPor: usuario['Usuário'],
    Arquivado: base.Arquivado || 'Não'
  };
}

function validarPaciente_(paciente) {
  if (!paciente.Nome) throw new Error('Nome é obrigatório.');
  if (!paciente.Telefone) throw new Error('Telefone é obrigatório.');
  if (!['Convênio', 'Particular'].includes(paciente.TipoPagamento)) throw new Error('Tipo de pagamento é obrigatório.');
  if (!STATUS_DISPONIVEIS.includes(paciente.Status)) throw new Error('Status inválido.');
  if (!paciente.PróximaAção) throw new Error('Próxima ação é obrigatória.');
  if (paciente.Status === 'Agendado' && (!paciente.DataCirurgia || !paciente.Hospital)) {
    throw new Error('Para status Agendado, informe data da cirurgia e hospital.');
  }
}

function pacientePrecisaAtencao_(paciente, configuracoes) {
  return motivosAtencao_(paciente, configuracoes).length > 0;
}

function motivosAtencao_(paciente, configuracoes) {
  const motivos = [];
  const hoje = hoje_();
  const statusFinal = ['Realizado', 'Cancelado'].includes(paciente.Status);
  const diasSemAtualizacao = diasEntre_(paciente.AtualizadoEm);
  const limiteSemAtualizacao = Number(configuracoes.DiasSemAtualizaçãoAlerta || 7);
  const limiteSolicitado = Number(configuracoes.DiasSolicitadoPlanoAlerta || 14);
  const limiteAutorizado = Number(configuracoes.DiasAutorizadoSemAgendamentoAlerta || 7);

  if (paciente.DataPróximoContato === hoje) motivos.push('Contato hoje');
  if (paciente.DataPróximoContato && paciente.DataPróximoContato < hoje) motivos.push('Contato atrasado');
  if (!statusFinal && !paciente.PróximaAção) motivos.push('Sem próxima ação');
  if (!statusFinal && diasSemAtualizacao !== null && diasSemAtualizacao > limiteSemAtualizacao) motivos.push('Sem atualização recente');
  if (paciente.Status === 'Solicitado ao plano de saúde' && diasSemAtualizacao !== null && diasSemAtualizacao > limiteSolicitado) motivos.push('Solicitação ao plano sem retorno');
  if (paciente.Status === 'Autorizado' && !paciente.DataCirurgia && diasSemAtualizacao !== null && diasSemAtualizacao > limiteAutorizado) motivos.push('Autorizado sem agendamento');
  if (paciente.Status === 'Pendente / travado') motivos.push('Pendente / travado');

  return motivos;
}

function aplicarValidacoesPacientes_() {
  const sheet = obterAba_('Pacientes');
  const hospitaisSheet = obterAba_('Hospitais');
  if (!sheet) return;
  const linhas = Math.max(sheet.getMaxRows() - 1, 1);
  const statusRule = SpreadsheetApp.newDataValidation().requireValueInList(STATUS_DISPONIVEIS, true).build();
  const pagamentoRule = SpreadsheetApp.newDataValidation().requireValueInList(['Convênio', 'Particular'], true).build();
  const proxRule = SpreadsheetApp.newDataValidation().requireValueInList(PROXIMAS_ACOES, true).build();
  const arquivadoRule = SpreadsheetApp.newDataValidation().requireValueInList(['Sim', 'Não'], true).build();
  const hospitalRule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(hospitaisSheet.getRange('B2:B'), true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(2, 7, linhas, 1).setDataValidation(pagamentoRule);
  sheet.getRange(2, 8, linhas, 1).setDataValidation(statusRule);
  sheet.getRange(2, 11, linhas, 1).setDataValidation(hospitalRule);
  sheet.getRange(2, 13, linhas, 1).setDataValidation(proxRule);
  sheet.getRange(2, 19, linhas, 1).setDataValidation(arquivadoRule);
}

// ===== Hospitais.gs =====
const HOSPITAIS_INICIAIS = [
  'Mater Dei Rio das Ostras',
  'Santa Helena',
  'São José',
  'Rede Casa'
];

function obterHospitais(sessaoOuUsuario) {
  sessaoOuUsuario && sessaoOuUsuario.Perfil ? sessaoOuUsuario : validarSessao_(sessaoOuUsuario);
  return lerRegistros_('Hospitais')
    .filter(item => item.Ativo === 'Sim')
    .sort((a, b) => String(a.Nome).localeCompare(String(b.Nome), 'pt-BR'));
}

function salvarHospital(nomeHospital, sessao) {
  const usuario = validarSessao_(sessao);
  verificarPermissao(usuario, 'cadastrarHospital');
  const nome = sanitizarTexto_(nomeHospital && nomeHospital.Nome ? nomeHospital.Nome : nomeHospital);
  if (!nome) throw new Error('Informe o nome do hospital.');

  const existente = lerRegistros_('Hospitais').find(item => String(item.Nome).toLowerCase() === nome.toLowerCase());
  if (existente) {
    if (existente.Ativo !== 'Sim') {
      existente.Ativo = 'Sim';
      atualizarLinha_('Hospitais', existente._linha, existente);
    }
    return criarResposta_(true, existente, 'Hospital já cadastrado.');
  }

  const hospital = { ID: gerarId_(), Nome: nome, CriadoEm: agora_(), CriadoPor: usuario['Usuário'], Ativo: 'Sim' };
  salvarLinha_('Hospitais', hospital);
  registrarHistorico({
    usuario: usuario,
    tipoAcao: 'Cadastro de hospital',
    valorNovo: nome,
    observacao: 'Hospital cadastrado.'
  });

  return criarResposta_(true, hospital, 'Hospital cadastrado.');
}

function semearHospitais_() {
  const existentes = lerRegistros_('Hospitais').map(item => String(item.Nome).toLowerCase());
  HOSPITAIS_INICIAIS.forEach(nome => {
    if (!existentes.includes(nome.toLowerCase())) {
      salvarLinha_('Hospitais', { ID: gerarId_(), Nome: nome, CriadoEm: agora_(), CriadoPor: 'Sistema', Ativo: 'Sim' });
    }
  });
}

// ===== WhatsApp.gs =====
const MODELOS_WHATSAPP_INICIAIS = [
  ['Andamento cirúrgico', 'Olá, tudo bem? Estamos entrando em contato para dar andamento à sua solicitação cirúrgica.'],
  ['Cobrar tomografia', 'Olá, tudo bem? Estamos aguardando o envio da sua tomografia para dar continuidade ao processo cirúrgico.'],
  ['Cirurgia autorizada', 'Olá, tudo bem? Sua cirurgia foi autorizada pelo plano. Podemos seguir com o agendamento?'],
  ['Confirmar cirurgia', 'Olá, tudo bem? Estamos entrando em contato para confirmar os dados da sua cirurgia.'],
  ['Cobrar exames', 'Olá, tudo bem? Passando para lembrar sobre os exames/documentos pendentes para sua cirurgia.']
];

function obterModelosWhatsApp(sessaoOuUsuario) {
  sessaoOuUsuario && sessaoOuUsuario.Perfil ? sessaoOuUsuario : validarSessao_(sessaoOuUsuario);
  return lerRegistros_('ModelosWhatsApp')
    .filter(item => item.Ativo === 'Sim')
    .sort((a, b) => String(a.Título).localeCompare(String(b.Título), 'pt-BR'));
}

function salvarModeloWhatsApp(dados, sessao) {
  const usuario = validarSessao_(sessao);
  verificarPermissao(usuario, 'editarModelosWhatsApp');

  const titulo = sanitizarTexto_(dados.Título || dados.titulo);
  const mensagem = sanitizarTexto_(dados.Mensagem || dados.mensagem);
  const ativo = sanitizarTexto_(dados.Ativo || dados.ativo || 'Sim');
  if (!titulo || !mensagem) throw new Error('Informe título e mensagem.');

  let modelo;
  if (dados.ID) {
    modelo = lerRegistros_('ModelosWhatsApp').find(item => item.ID === dados.ID);
  }

  if (modelo) {
    const anterior = Object.assign({}, modelo);
    modelo.Título = titulo;
    modelo.Mensagem = mensagem;
    modelo.Ativo = ativo === 'Não' ? 'Não' : 'Sim';
    atualizarLinha_('ModelosWhatsApp', modelo._linha, modelo);
    registrarHistorico({
      usuario: usuario,
      tipoAcao: 'Edição de modelo WhatsApp',
      campoAlterado: 'Modelo',
      valorAnterior: anterior.Título,
      valorNovo: titulo,
      observacao: 'Modelo de WhatsApp atualizado.'
    });
  } else {
    modelo = { ID: gerarId_(), Título: titulo, Mensagem: mensagem, Ativo: ativo === 'Não' ? 'Não' : 'Sim', CriadoEm: agora_() };
    salvarLinha_('ModelosWhatsApp', modelo);
    registrarHistorico({
      usuario: usuario,
      tipoAcao: 'Cadastro de modelo WhatsApp',
      valorNovo: titulo,
      observacao: 'Modelo de WhatsApp cadastrado.'
    });
  }

  return criarResposta_(true, modelo, 'Modelo salvo.');
}

function formatarTelefoneWhatsApp(telefone) {
  let numero = String(telefone || '').replace(/\D/g, '');
  if (!numero) return '';
  if (numero.startsWith('55')) return numero;
  if (numero.length === 10 || numero.length === 11) return '55' + numero;
  return numero;
}

function gerarLinkWhatsApp(telefone, mensagem, idPaciente, sessao) {
  const usuario = validarSessao_(sessao);
  verificarPermissao(usuario, 'usarWhatsApp');
  const numero = formatarTelefoneWhatsApp(telefone);
  if (!numero) throw new Error('Telefone inválido para WhatsApp.');
  const link = 'https://wa.me/' + numero + '?text=' + encodeURIComponent(mensagem || '');

  let paciente = null;
  if (idPaciente) paciente = encontrarPaciente_(idPaciente);

  registrarHistorico({
    usuario: usuario,
    idPaciente: paciente ? paciente.ID : '',
    nomePaciente: paciente ? paciente.Nome : '',
    tipoAcao: 'WhatsApp',
    observacao: 'Contato por WhatsApp iniciado.'
  });

  return criarResposta_(true, { link: link }, 'Link gerado.');
}

function semearModelosWhatsApp_() {
  const existentes = lerRegistros_('ModelosWhatsApp').map(item => String(item.Título).toLowerCase());
  MODELOS_WHATSAPP_INICIAIS.forEach(item => {
    if (!existentes.includes(item[0].toLowerCase())) {
      salvarLinha_('ModelosWhatsApp', { ID: gerarId_(), Título: item[0], Mensagem: item[1], Ativo: 'Sim', CriadoEm: agora_() });
    }
  });
}

// ===== PlanosSaude.gs =====
function obterPlanosSaude(sessaoOuUsuario) {
  sessaoOuUsuario && sessaoOuUsuario.Perfil ? sessaoOuUsuario : validarSessao_(sessaoOuUsuario);
  return lerRegistros_('PlanosDeSaúde')
    .filter(item => item.Ativo === 'Sim')
    .sort((a, b) => String(a.Nome).localeCompare(String(b.Nome), 'pt-BR'));
}

function salvarPlanoSaude(nomePlano, sessao) {
  const usuario = validarSessao_(sessao);
  verificarPermissao(usuario, 'cadastrarPlano');
  const nome = sanitizarTexto_(nomePlano && nomePlano.Nome ? nomePlano.Nome : nomePlano);
  if (!nome) throw new Error('Informe o nome do plano de saúde.');
  return criarResposta_(true, salvarPlanoSeNovo_(nome, usuario), 'Plano de saúde salvo.');
}

function salvarPlanoSeNovo_(nomePlano, usuario) {
  const nome = sanitizarTexto_(nomePlano);
  if (!nome) return null;
  const existente = lerRegistros_('PlanosDeSaúde').find(item => String(item.Nome).toLowerCase() === nome.toLowerCase());
  if (existente) return existente;

  const plano = { ID: gerarId_(), Nome: nome, Ativo: 'Sim', CriadoEm: agora_(), CriadoPor: usuario['Usuário'] || 'Sistema' };
  salvarLinha_('PlanosDeSaúde', plano);
  registrarHistorico({
    usuario: usuario,
    tipoAcao: 'Cadastro de plano de saúde',
    valorNovo: nome,
    observacao: 'Plano de saúde cadastrado.'
  });
  return plano;
}
