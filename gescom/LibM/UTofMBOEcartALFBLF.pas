unit UTofMBOEcartALFBLF;

interface
uses StdCtrls, Controls, Classes, forms, sysutils, ComCtrls, UTOF, EntGc, UTob,
  {$IFDEF EAGLCLIENT}
  eQRS1,
  {$ELSE}
  QRS1, EdtEtat, DBTables,
  {$ENDIF}
  hmsgbox, M3FP, HCtrls, HEnt1, FactComm, factutil, CalcOLEGescom,
  HQry ;

type
  TOF_MBOEcartALFBLF = class(TOF)
    procedure OnArgument(stArgument: string); override;
    procedure OnClose; override;
    procedure LanceEdition;
    procedure ChangeCritere;

  end;

implementation

procedure TOF_MBOEcartALFBLF.OnArgument(stArgument: string);
begin
  inherited;
  // paramétrage du libellé etablissement quand on est en multi-dépôts
  if VH_GC.GCMultiDepots then
  begin
   SetControlText('TGP_DEPOT', TraduireMemoire('Dépôt émetteur'));
   SetControlProperty('DEPOT','DataType','GCDEPOT');
  end
  else SetControlText('TGP_DEPOT', TraduireMemoire('Etablissement émetteur'));
  if stArgument = 'ALF' then
    Ecran.Caption := TraduireMemoire('Edition des écarts entre annonce de livraison et réception');
  if stArgument = 'CF' then
    Ecran.Caption := TraduireMemoire('Edition des écarts entre commande fournisseur et réception');
  updatecaption(Ecran);
end;

procedure TOF_MBOEcartALFBLF.OnClose;
begin
  // Suppression des enregistrements concernant cet utilisateur
  ExecuteSQL('DELETE FROM GCTMPECART WHERE GZT_UTILISATEUR = "' + V_PGI.USer + '"');
  VH_GC.TOBEdt.ClearDetail;
end;

procedure TOF_MBOEcartALFBLF.ChangeCritere;
begin
  VH_GC.TOBEdt.ClearDetail;
end;

// Procedure de lancement de l'état
procedure TOF_MBOEcartALFBLF.LanceEdition;
var
  stSQL, stwhere, stWhereNumero, Nature: string;
  F: TFQRS1;
  Q : TQuery;
  TOBALF : TOB;
  i: Integer;
  Edit: THEdit;
  Edit2: THValcombobox;
begin
  F := TFQRS1(Ecran);
  Nature := THEdit(TForm(Ecran).FindComponent('NATURE')).Text;
  // Suppression des enregistrements concernant cet utilisateur
  ExecuteSQL('DELETE FROM GCTMPECART WHERE GZT_UTILISATEUR = "' + V_PGI.USer + '"');

  //***************************** CONSTRUCTION DU WHERE ************************
  // Le critère DEPOT
  Edit2 := THValComboBox(TFQRS1(F).FindComponent('DEPOT'));
  if (Edit2 <> nil) and (Edit2.Value <> '') and (Edit2.Value <> TraduireMemoire('<<Tous>>')) then
  begin
    if stWhere <> '' then stWhere := stWhere + ' AND ';
    stWhere := 'GP_DEPOT = "' + Edit2.Value + '"';
  end;

  // Le critère TIERS
  Edit := THEdit(TFQRS1(F).FindComponent('TIERS'));
  if (Edit <> nil) and (Edit.Text <> '') then
  begin
    if stWhere <> '' then stWhere := stWhere + ' AND ';
    stWhere := stWhere + 'GP_TIERS = "' + Edit.Text + '"';
  end;

  // ********************************************************
  //Le critère CODE ARTICLE
  Edit := THEdit(TFQRS1(F).FindComponent('REFARTSAISIE'));
  if (Edit <> nil) and (Edit.Text <> '') then
  begin
    // pour articles dimensionnés et uniques
    if stWhere <> '' then stWhere := stWhere + ' AND ';
    stWhere := stWhere + 'GL_ARTICLE like "' + Edit.Text + '%"';
  end;

  // ********************************************************
  //Le critère DATEPIECE
  Edit := THEdit(TFQRS1(F).FindComponent('DATEPIECE'));
  if (Edit <> nil) and (Edit.Text <> '') then
  begin
    if stWhere <> '' then stWhere := stWhere + ' AND ';
    stWhere := stWhere + 'GP_DATEPIECE >= "' + USDateTime(StrToDate(Edit.Text)) + '"';
  end;

  // Le critère NUMERO et NUMERO_
  Edit := THEdit(TFQRS1(F).FindComponent('NUMERO'));
  if (Edit <> nil) and (Edit.Text <> '') then
  begin
    if stWhereNumero <> '' then stWhereNumero := stWhereNumero + ' AND ';
    stWhereNumero := stWhereNumero + 'GP_NUMERO >= "' + Edit.Text + '"';
  end;
  Edit := THEdit(TFQRS1(F).FindComponent('NUMERO_'));
  if (Edit <> nil) and (Edit.Text <> '') then
  begin
    if stWhereNumero <> '' then stWhereNumero := stWhereNumero + ' AND ';
    stWhereNumero := stWhereNumero + 'GP_NUMERO <= "' + Edit.Text + '"';
  end;
  //****************************************************************************

  stSQL := 'SELECT "' + V_PGI.USer + '" AS GZT_UTILISATEUR, ' +
    'L1.GL_NATUREPIECEG||";"||L1.GL_SOUCHE||";"||trim(str(L1.GL_NUMERO))||";"||trim(STR(L1.GL_INDICEG)) AS GZT_EMETTEUR, ' +
    'L1.GL_NUMLIGNE AS GZT_NUMLIGEM, '+
    'L1.GL_ARTICLE AS GZT_ARTICLE, GP_DEPOT AS GZT_DEPOTEM, GP_DEPOTDEST AS GZT_DEPOTRE, ' +
    'L1.GL_TIERS AS GZT_TIERS, L1.GL_DEVISE AS GZT_DEVISE, L1.GL_DATEPIECE AS GZT_DATEEM, ' +
    'SUM(L1.GL_QTEFACT) AS GZT_QTEEMISE, GP_NATUREPIECEG, GP_NUMERO ' +
    'FROM PIECE LEFT JOIN LIGNE L1 ON L1.GL_NATUREPIECEG=GP_NATUREPIECEG AND L1.GL_SOUCHE=GP_SOUCHE ' +
    'AND L1.GL_NUMERO=GP_NUMERO AND L1.GL_INDICEG=GP_INDICEG ' +
    'WHERE gp_naturepieceg="' + Nature + '"';
  if stWhereNumero <> '' then stSQL := stSQL + ' and ' + stWhereNumero;
  if stWhere <> '' then stSQL := stSQL + ' and ' + stWhere;
  stSQL := stSQL +  ' AND GP_INDICEG=0 AND GP_SUPPRIME<>"X" AND L1.GL_ARTICLE<>"" ' +
    'GROUP BY L1.GL_NATUREPIECEG||";"||L1.GL_SOUCHE||";"||trim(str(L1.GL_NUMERO))||";"||trim(STR(L1.GL_INDICEG)),' +
    'L1.GL_NUMLIGNE, L1.GL_ARTICLE, GP_DEPOT, GP_DEPOTDEST, L1.GL_TIERS, L1.GL_DEVISE, L1.GL_DATEPIECE, ' +
    'GP_NATUREPIECEG, GP_NUMERO ' +
    'ORDER BY GZT_EMETTEUR,GZT_ARTICLE';

  // Si la TOB n'existe pas ou qu'il y a aucun enregistrement alors
  // je la crée et la remplie
  TOBALF := TOB.Create('', nil, -1);
  Q := OpenSQL(stSQL, True);
  TOBALF.LoadDetailDB('', '', '', Q, False);
  Ferme(Q);
  For i := 0 to TOBALF.Detail.count - 1 do
  begin
    stSQL := 'SELECT SUM(GL_QTEFACT) AS GZT_QTERECUE ' +
      'FROM LIGNE ' +
      'WHERE GL_NATUREPIECEG="BLF" ' +
      'AND GL_ARTICLE = "' + TOBALF.Detail[i].GetValue( 'GZT_ARTICLE' ) + '" ' +
      'AND GL_PIECEPRECEDENTE LIKE "%;'+TOBALF.Detail[i].GetValue( 'GP_NATUREPIECEG' )+
        ';___;'+IntToStr(TOBALF.Detail[i].GetValue( 'GP_NUMERO' ))+';%" ';
    Q := OpenSQL(stSQL, True);
    TOBALF.Detail[i].AddChampSup( 'GZT_QTERECUE', False );
    TOBALF.Detail[i].PutValue( 'GZT_QTERECUE', Q.FindField('GZT_QTERECUE').AsString );
    TOBALF.Detail[i].AddChampSup( 'GZT_COMPTEUR', False );
    TOBALF.Detail[i].PutValue( 'GZT_COMPTEUR', i);
    TOBALF.Detail[i].VirtuelleToReelle( 'GCTMPECART' );
  end;
  TOBALF.InsertDB( nil );
  // Clause where afin d'être sûre de récupérer ses propres enregistrements
  THEdit(TForm(Ecran).FindComponent('GZT_UTILISATEUR')).Text := V_PGI.USer;
end;

// Procedure AGL pour lancer le traitement pour l'édition des écarts

procedure AGLOnClickLanceEdition(parms: array of variant; nb: integer);
var F: TForm;
  TOTOF: TOF;
begin
  F := TForm(Longint(Parms[0]));
  if (F is TFQRS1) then TOTOF := TFQRS1(F).LaTOF
  else exit;
  if (TOTOF is TOF_MBOEcartALFBLF) then TOF_MBOEcartALFBLF(TOTOF).LanceEdition;
end;

procedure AGLOnChangeCritereALFBLF(parms: array of variant; nb: integer);
var F: TForm;
  TOTOF: TOF;
begin
  F := TForm(Longint(Parms[0]));
  if (F is TFQRS1) then TOTOF := TFQRS1(F).LaTOF
  else exit;
  if (TOTOF is TOF_MBOECARTALFBLF) then TOF_MBOECARTALFBLF(TOTOF).ChangeCritere;
end;

initialization
  registerclasses([TOF_MBOEcartALFBLF]);
  RegisterAglProc('OnChangeCritereALFBLF', TRUE, 1, AGLOnChangeCritereALFBLF);
  RegisterAglProc('OnClickLanceEdition', TRUE, 1, AGLOnClickLanceEdition);
end.

