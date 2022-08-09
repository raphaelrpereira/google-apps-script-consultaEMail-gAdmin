/**
 * Conecta com banco de dados na Hostgator onde estão os dados dos alunos que estão sem e-mail institucional no RM
 */
function integracao() {

  var GUIA = "Dados RM"; // Aba para imprimir e consultar os dados importados do banco de dados MySQL hospedado na Hostgator
  var start = new Date(); // Debug, para saber o tempo de execução do script

  var doc = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = doc.getSheetByName(GUIA);
  var lastRow = sheet.getLastRow();
  var cell = sheet.getRange(lastRow+1,1);

  //formatando coluna D da planilha para receber a informação como uma String (CODCURSO)
  doc.getRange('D:D').setNumberFormat('@');

  //obtendo dados que atualmente estão na planilha
  var dadosAtuais = getMatriculasNaPlanilha(GUIA);

  // Conexão com o banco de dados e execução da query
  var conn = conectaBanco();
  var query = conn.createStatement();
  var result = query.executeQuery("SELECT * FROM emails_institucionais WHERE email_educacional IS NULL");
  
  var row = 0;
  while(result.next()){

    //verifica se RA já existe na planilha. Só insere se RA não existir.
    var ra = result.getString(8);

    // se não existir a matrícula na planilha faz a inserção
    if( !dadosAtuais.includes(ra) ){
        Logger.log('Cadastrando RA ' + ra);
        
        for (var col = 0; col < result.getMetaData().getColumnCount(); col++) { 
          cell.offset(row, col).setValue(result.getString(col + 1)); 
        }
        row++;
    }
    
  }

  //buscando e-mails institucionais de cada aluno na planilha
  buscaEmailsGoogle(GUIA);
  
  //redimensionando colunas de acordo com o conteúdo preenchido
  doc.getActiveSheet().autoResizeColumns(1,14);
  
  result.close();
  query.close();
  conn.close();
  var end = new Date();
  
  // Geramos um log de tempo execução
  Logger.log('Última execução em: ' + start + '. Tempo de execução: ' + ( (end.getTime() - start.getTime() ) / 1000) + ' segundos'); 
}

function conectaBanco(){
    var BANCO = "mysql"; // Conector do banco
    var HOST = "HOSTNAME"; 
    var PORT = "3306"; 
    var DATABASE = "DATABASENAME" 
    var USER = "USERNAME"; 
    var PASSWORD = "PASSWORD"; 
    return Jdbc.getConnection("jdbc:" + BANCO +"://" + HOST + ":" + PORT + "/" + DATABASE,  USER, PASSWORD);
}


/**
 *    Alimenta Array com matrículas que já estão na planilha no momento do início da execução
 */
function getMatriculasNaPlanilha(guia){

    var dados = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(guia);
    var maxRows = dados.getLastRow();    

    matriculas = new Array(maxRows);
    
    // percorre todas as linhas já preenchidas para verificar se o valor já existe.
    for( var linha = 1; linha <= maxRows; linha++ ){
      matriculas[linha] = dados.getRange(linha, 8).getValue();
    }

    return matriculas;   
}

/**
 *    Realiza a busca pelos e-mails institucionais no Admin Google para cada linha da planilha.
 */
function buscaEmailsGoogle(sheetName){
  
  var planilha = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  var maxRows = planilha.getLastRow();

  for( row = 2; row <= maxRows; row++){
    
    //se campo e-mail institucional está vazio faz a busca pelo nome, do contrário sai da função.
    if( planilha.getRange(row,11).getValue() == ""  ){

        var nome    =   planilha.getRange(row,7).getValue();
        var emailEducacional   =   listUsersContains(nome);
        
        if(emailEducacional){
          
           var ra           =   planilha.getRange(row,8).getValue();
           var emailPessoal =   planilha.getRange(row,10).getValue();
           var campus       =   planilha.getRange(row,2).getValue();
           
           // se existe e-mail, insere na planilha
           planilha.getRange(row,11).setValue(emailEducacional);
           planilha.getRange(row,13).setValue(new Date());

           // e insere no banco de dados mysql
           updateEmailsDB( ra, emailEducacional);

           // em seguida, se o e-mail pessoal não estiver vazio cria o documento e envia por e-mail .
           if(emailPessoal != ""){
                sendMailEducacional( emailPessoal, emailEducacional, nome, campus);
           }else{
                notificaSemEmailPessoal(emailEducacional, nome, campus);
           }
           

        }

    }

  }

} 

/**
 *    Realiza a busca do e-mail institucional no Google Admin pelo NOME recebido no parâmetro.
 */
function listUsersContains(nomeAluno) {
  let pageToken;
  let page;
  
  var splitName = nomeAluno.split(' ');
  var firstName = splitName[0];

  do {
    page = AdminDirectory.Users.list({
      domain: 'subdomain.ies.edu.br',
      "query": "givenName:" + firstName,
      orderBy: 'givenName',
      maxResults: 2,
      pageToken: pageToken
    });

    const users = page.users;
    
    if (!users) {
      //Logger.log('No users found.');
      return;
    }
    
    for (const user of users) {
      
      //remove acentuação do fullName para comparar com nomeAluno
      var fullName2 = removeAcentos(user.name.fullName);

      // Verifica se o nome completo do usuário é o nome completo enviado 
      if( fullName2 == nomeAluno ){
      
        return user.primaryEmail;
      
      }      
    
    }
    pageToken = page.nextPageToken;

  } while (pageToken);

}

/**
 * Retorna a string sem acentuação.
 */
function removeAcentos(nome){   
  return nome.normalize('NFD').replace(/[\u0300-\u036f]/g, "");  
}


/**
 *    Atualiza e-mail no banco de dados MySQL --> Hostgator
 */
function updateEmailsDB(ra, email){

  try{
        var conn = conectaBanco();
        var query = conn.prepareStatement('UPDATE emails_institucionais SET email_educacional = ?, atualizado_em = NOW() WHERE ra = ?');
        query.setString(1, email);
        query.setString(2, ra);
        query.execute();
  }catch(err){
    Logger.log('Falha com o erro %s', err.message);
  
  }
 
}



/**
 *    Gera Comunicado de Boas vindas e envia para o e-mail pessoal do aluno.
 */
function sendMailEducacional(emailPessoal, emailEducacional, nome, campus){
    
    // pega o ID do documento Google que será utilizado como modelo para a criação de um novo (mala direta)
    var idDocAraguari = 'idDocumentoGoogle1';
    var idDocItumbara = 'IdDocumentoGoogle2';
    
    idDoc = (campus == 'Araguari') ? idDocAraguari : idDocItumbara;

    // informações do remetente e destinatario 
    var nome_completo = nome;
    var destinatario = emailPessoal;
    var subject = "Sua conta Google IES foi criada";
    var body = " Olá " + nome + ". \n Você está recebendo em anexo o acesso ao seu e-mail educacional. \n Este é um e-mail de disparo automático. Quaisquer dúvidas entre em contato conosco pelo e-mail suporte@nomeies.edu.br.";
    var remetente = "NOME <suporte@nomeies.edu.br>";

    // Cria um documento temporário, recupera o ID e o abre
    var idCopia = DriveApp.getFileById(idDoc).makeCopy('Acesso Email Institucional ' + nome).getId();
    var docCopia = DocumentApp.openById(idCopia);

    // recupera o corpo do documento
    var bodyCopia = docCopia.getActiveSection();

    // faz o replace das variáveis do template, salva e fecha o documento temporario
    bodyCopia.replaceText("<<nome_aluno>>", nome_completo);
    bodyCopia.replaceText("<<email_educacional>>", emailEducacional);
    docCopia.saveAndClose();

    // abre o documento temporario como PDF utilizando o seu ID
    var pdf = DriveApp.getFileById(idCopia).getAs("application/pdf");

    // envia o email
    MailApp.sendEmail(destinatario, subject, body, {name: remetente, attachments: pdf, bcc:'teste@nomeies.edu.br'});

    // apaga o documento temporário
    DriveApp.getFileById(idCopia).setTrashed(true);

    Logger.log('Informações de acesso enviadas para ' + nome + ' - ' + emailPessoal);

}

function notificaSemEmailPessoal(emailEducacional, nome, campus){
    var idDocAraguari = '1FEI65Qfks5FYM-vno4NYMkH8Q2Bko_p_a2kcw-nsOEA';
    var idDocItumbara = '1QWkLch5QS7JQ31meu6l_kToelA9xO4uA52B9U9Yt2ek';
    
    idDoc = (campus == 'Araguari') ? idDocAraguari : idDocItumbara;

    // informações do remetente e destinatario 
    var nome_completo = nome;
    var destinatario = 'suporte@nomeies.edu.br';
    var subject = "Conta Google do aluno " + nome + " criada.";
    var body = " Entrando em contato para informar que a conta Google do aluno " + nome + " foi criada mas não foi possível enviar a notificação para o aluno. \n Motivo: E-mail pessoal não cadastrado. ";
    var remetente = "NOMEIES <suporte@nomeies.edu.br>";

    // Cria um documento temporário, recupera o ID e o abre
    var idCopia = DriveApp.getFileById(idDoc).makeCopy('Acesso Email Institucional ' + nome).getId();
    var docCopia = DocumentApp.openById(idCopia);

    // recupera o corpo do documento
    var bodyCopia = docCopia.getActiveSection();

    // faz o replace das variáveis do template, salva e fecha o documento temporario
    bodyCopia.replaceText("<<nome_aluno>>", nome_completo);
    bodyCopia.replaceText("<<email_educacional>>", emailEducacional);
    docCopia.saveAndClose();

    // abre o documento temporario como PDF utilizando o seu ID
    var pdf = DriveApp.getFileById(idCopia).getAs("application/pdf");

    // envia o email
    MailApp.sendEmail(destinatario, subject, body, {name: remetente, attachments: pdf, bcc:'test@nomeies.edu.br'});

    // apaga o documento temporário
    DriveApp.getFileById(idCopia).setTrashed(true);

    Logger.log('Informações de acesso enviadas para ' + nome + ' - ' + emailPessoal);


}
