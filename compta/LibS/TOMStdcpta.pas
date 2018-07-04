unit TOMStdcpta;

interface
uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  assist, HSysMenu, hmsgbox, StdCtrls, HTB97, ComCtrls, ExtCtrls, Hctrls,
  ActnList, Mask, HPanel, UTOB, HQuickRP, dbtables, ImgList, Menus,
  HEnt1,LicUtil,MajTable, PGIExec, ParamSoc,fe_main, HStatus, Fiche,
  FichComm,Ent1, Devise, MulJal, MulGene, ParamLib, ContAbon, Guide, Exercice, AssistExo,
  Spin,galOutil,UTOM,Grids,DBGrids,HDB,Mul,Fichlist,
  DBCtrls,Vierge, uLibStdCpta;


  procedure SupprimeDonneStd (TableStd : string; numstd : integer);
  procedure SupprimeDonneRefStd (TableStd : string);
  procedure SuppressionPlanNonreference;

Type
     TOM_STDCPTA = Class (TOM)
       procedure OnArgument (stArgument : String ) ; override ;
       procedure DbClickDelete(Sender: TObject);
       procedure OnLoadRecord; override;
       procedure OnNewRecord; override;
       procedure OnClose; override;
       procedure FormKeyDownOrga(Sender: TObject; var Key: Word; Shift: TShiftState);
       private
       procedure ChargeParamSocRef (NumPlan : integer; Charge : Boolean);
       procedure bFinClick(Sender: TObject);
       procedure ChargeStandardCompta;
       procedure ChargeAxe;
       function  ChargeStandard : boolean;
       procedure SupprimeTableChoixCod(Code : string);
       procedure SupprimeTableStd (TableDos : string);
       procedure LoadStandardRef (NumStd : integer; TableDos,TableStd : string);
       procedure SupprimeTableChoixListe(LIST : string);
     END ;
var
NumPlanCompte    : integer;

implementation

//uses AssistPl;

procedure TOM_STDCPTA.OnArgument(stArgument: String);
begin
  THDBGRID(GetControl('FListe')).OnDblclick := bFinClick;
  TToolbarButton97(GetControl('BDelete')).Onclick := DbClickDelete;
  Ecran.OnKeyDown :=  FormKeyDownOrga;
//  TToolbarButton97(GetControl('BValider')).Onclick := bFinClick;
end;
procedure TOM_STDCPTA.OnLoadRecord;
var
Q1        : TQuery;
begin

  if GetParamSoc ('SO_LGCPTEGEN')= 0 then
     SetParamSoc ('SO_LGCPTEGEN', 6);
  if GetParamSoc ('SO_LGCPTEAUX')= 0 then
     SetParamSoc ('SO_LGCPTEAUX', 6);

  if GetParamSoc ('SO_BOURREGEN') <> '0' then
     SetParamSoc ('SO_BOURREGEN', '0');
  if GetParamSoc ('SO_BOURREAUX') <> '0' then
     SetParamSoc ('SO_BOURREAUX', '0');

      // Chargement VH
  ChargeSocieteHalley;

  Q1 := OpenSQl ('SELECT * From SOCIETE', TRUE);
  if Q1.EOF then InitSocietePCL ;
  Ferme (Q1);

  Q1 := OpenSQl ('SELECT * From ETABLISS', TRUE);
  if Q1.EOF then InitEtablissement;
  Ferme (Q1);

end;
procedure TOM_STDCPTA.OnClose;
begin
if NumPLanCompte = 0 then
     NumPlanCompte := GetParamSoc('SO_NUMPLANREF');
end;

procedure TOM_STDCPTA.FormKeyDownOrga(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
Case (ORD(Key)) of
     VK_F5 :     // double click
               begin Key := 0; bFinClick(Sender);  end;
     VK_DELETE : // suppression
               DbClickDelete(Sender);
     VK_ESCAPE   : Key := 0;
     VK_F10    :  begin Key := 0;  TFFiche(Ecran).BValiderClick(nil); end;
end;

end;


procedure TOM_STDCPTA.OnNewRecord;
begin
SetParamSoc ('SO_NUMPLANREF', '0');
end;

procedure TOM_STDCPTA.DbClickDelete(Sender: TObject);
var
Texte    : string;
OM       : TOM;
NoStd    : integer;
begin
if not IsSuperviseur(TRUE) then
begin
PGIInfo('Vous n''êtes pas un utilisateur privilégié','') ; exit;
end;

NoStd :=  GetField('STC_NUMPLAN');
if not EstSpecif('51502') then
if noStd < 21 then
begin PGIInfo('Plan type CEGID,impossible de supprimer','') ;  exit;end;

if not (TFFicheListe(Ecran).Bouge(nbDelete)) then exit;
SupprimeDonneStd ('PARSOCREF',NoStd);
SupprimeDonneStd ('GENERAUXREF',NoStd);
SupprimeDonneStd ('JALREF',NoStd);
SupprimeDonneStd ('GUIDEREF',NoStd);
SupprimeDonneStd ('ECRGUIREF',NoStd);
SupprimeDonneStd ('ANAGUIREF',NoStd);
SupprimeDonneStd ('TIERSREF',NoStd);
SupprimeDonneStd ('CORRESPREF',NoStd);
SupprimeDonneStd ('RUPTUREREF',NoStd);
SupprimeDonneStd ('REFAUTOREF',NoStd);
SupprimeDonneStd ('AXEREF',NoStd);
SupprimeDonneStd ('SECTIONREF',NoStd);
SupprimeDonneStd ('CHOIXCODREF',NoStd);
SupprimeDonneStd ('VENTILREF',NoStd);
SupprimeDonneStd ('NATCPTEREF',NoStd);
SupprimeDonneStd ('FILTRESREF',NoStd);
SupprimeDonneStd ('LISTEREF',NoStd);
// CA - 02/04/2002
SupprimeDonneStd ('MODEPAIEREF',NoStd);
SupprimeDonneStd ('MODEREGLREF',NoStd);
// Fin CA - 02/04/2002

end;

procedure SupprimeDonneStd (TableStd : string; numstd : integer);
var QDos : TQuery;
PrefStd  : String;
Where    : string;
begin
  PrefStd := TableToPrefixe(TableStd);

Where := ' WHERE '+PrefStd+'_NUMPLAN='+IntToStr(numstd);

  QDos := OpenSQL('SELECT * FROM '+TableStd + Where,True);
  if not QDos.Eof then
      ExecuteSQL('Delete from '+TableStd + Where);
  Ferme (QDos);
end;

procedure SupprimeDonneRefStd (TableStd : string);
var QDos  : TQuery;
PrefStd   : String;
Predefini : string;
Where     : string;
begin
  PrefStd := TableToPrefixe(TableStd);

  Where := ' WHERE '+PrefStd+'_NUMPLAN < 21 AND '+ PrefStd+'_PREDEFINI="STD"';

  QDos := OpenSQL('SELECT * FROM '+TableStd + Where,True);
  if not QDos.Eof then
      ExecuteSQL('Delete from '+TableStd + Where);
  Ferme (QDos);

  Where := ' WHERE '+PrefStd+'_NUMPLAN >= 100';
  QDos := OpenSQL('SELECT * FROM '+TableStd + Where,True);
  if not QDos.Eof then
      ExecuteSQL('Delete from '+TableStd + Where);

  Ferme (QDos);
end;

procedure SuppressionPlanNonreference;
var
Q1        : TQuery;
Where     : string;
begin
  // pour supprimer tous les standards predefinis STD et < 21
  // et tous les standards > 100
  SupprimeDonneRefStd ('PARSOCREF');
  SupprimeDonneRefStd ('GENERAUXREF');
  SupprimeDonneRefStd ('JALREF');
  SupprimeDonneRefStd ('GUIDEREF');
  SupprimeDonneRefStd ('ECRGUIREF');
  SupprimeDonneRefStd ('ANAGUIREF');
  SupprimeDonneRefStd ('TIERSREF');
  SupprimeDonneRefStd ('CORRESPREF');
  SupprimeDonneRefStd ('RUPTUREREF');
  SupprimeDonneRefStd ('REFAUTOREF');
  SupprimeDonneRefStd ('AXEREF');
  SupprimeDonneRefStd ('SECTIONREF');
  SupprimeDonneRefStd ('CHOIXCODREF');
  SupprimeDonneRefStd ('VENTILREF');
  SupprimeDonneRefStd ('NATCPTEREF');
  SupprimeDonneRefStd ('FILTRESREF');
  SupprimeDonneRefStd ('LISTEREF');
// CA - 02/04/2002
SupprimeDonneRefStd ('MODEPAIEREF');
SupprimeDonneRefStd ('MODEREGLREF');
// Fin CA - 02/04/2002

  Where := ' WHERE STC_NUMPLAN < 21 AND STC_PREDEFINI="STD"';
  Q1 := OpenSQL('SELECT * FROM STDCPTA '+ Where,True);
  if not Q1.Eof then
      ExecuteSQL('Delete from STDCPTA' + Where);
  Ferme (Q1);

  Where := ' WHERE STC_NUMPLAN >= 100';
  Q1 := OpenSQL('SELECT * FROM STDCPTA '+ Where,True);
  if not Q1.Eof then
      ExecuteSQL('Delete from STDCPTA' + Where);
  Ferme (Q1);

end;

procedure TOM_STDCPTA.ChargeParamSocRef (NumPlan : integer; Charge : Boolean);
var QParamSocRef : TQuery ;
    StVal : string;
    lgene, lgaux    : integer;
begin
  QParamSocRef := OpenSQL ('SELECT * FROM PARSOCREF WHERE PRR_NUMPLAN = ' +
                      IntToStr(NumPlan) + ' ORDER BY PRR_SOCNOM',True) ;
  InitMove(QCount(QParamSocRef),'Chargement des paramètres société');
  while not QParamSocRef.EOF do
  begin
    StVal := QParamSocRef.FindField('PRR_SOCDATA').AsString;
    if (QParamSocRef.FindField('PRR_COMPTE').AsString='G') then
       StVal := BourreLaDonc(stVal,fbGene)
    else if (QParamSocRef.FindField('PRR_COMPTE').AsString='T') then
       StVal := BourreLaDonc(stVal,fbAux);
    if QParamSocRef.FindField('PRR_SOCNOM').AsString = 'SO_LGCPTEGEN' then
    begin
    if Charge then
        lgene :=  StrToInt(QParamSocRef.FindField('PRR_SOCDATA').AsString);
        SetParamSoc ('SO_LGCPTEGEN', lgene);
    end;
    if QParamSocRef.FindField('PRR_SOCNOM').AsString = 'SO_LGCPTEAUX' then
    begin
    if Charge then
       lgaux :=  StrToInt(QParamSocRef.FindField('PRR_SOCDATA').AsString);
       SetParamSoc ('SO_LGCPTEAUX', lgaux);
    end;
    SetParamSoc ( QParamSocRef.FindField('PRR_SOCNOM').AsString, StVal);
    QParamSocRef.Next ;
    MoveCur (False);
  end;
  Ferme(QParamSocRef);
  FiniMove;
end;

procedure TOM_STDCPTA.bFinClick(Sender: TObject);
var
ValChamp,where        : string;
Q1                    : Tquery;
Texte                 : string;
begin
  inherited;
//    NumPlanCompte := GetField('STC_NUMPLAN');

    NumPlanCompte := TSpinEdit(GetControl('STC_NUMPLAN')).value;
    ChargeParamSocRef(NumPlanCompte, TRUE);

    Texte := 'Voulez-vous recharger le Plan ' + IntToStr (NumPlanCompte);
    if HShowMessage('0;Suppression des paramètres;'+Texte+';Q;YN;N;N;','','')<>mrYes then exit ;

    SupprimeTableStd ('GENERAUX');
    SupprimeTableStd ('TIERS');
    SupprimeTableStd ('JOURNAL');
    SupprimeTableStd ('GUIDE');
    SupprimeTableStd ('ECRGUI');
    SupprimeTableStd ('ANAGUI');
    SupprimeTableStd ('CORRESP');
    SupprimeTableStd ('RUPTURE');
    SupprimeTableStd ('REFAUTO');
    SupprimeTableStd ('AXE');
    SupprimeTableStd ('SECTION');
    SupprimeTableChoixCod ('VTY');
    SupprimeTableChoixCod ('RUG');
    SupprimeTableChoixCod ('RUT');
    SupprimeTableStd ('VENTIL');
    SupprimeTableStd ('NATCPTE');
    SupprimeTableStd ('FILTRES'); // CA - 28/11/2001   
    SupprimeTableChoixListe ('MULMMVTS');
    SupprimeTableChoixListe ('MULMANAL');
    SupprimeTableChoixListe ('MULVMVTS');

    SetParamSoc ('SO_NUMPLANREF', NumPlanCompte);
    SetParamSoc ('SO_BOURREGEN', '0');
    SetParamSoc ('SO_BOURREAUX', '0');

  if (NumPlanCompte > 0 ) and (NumPlanCompte <> GetParamSoc('SO_NUMPLANREF'))then
     SetParamSoc ('SO_NUMPLANREF', NumPlanCompte);
  ChargeMagHalley;
  ChargeStandard;
  Ecran.ModalResult := MROK;
  RajouteCaptionDossier('Dossier type : '+ IntToStr (NumPlanCompte));

//  THPanel(TFVierge(Ecran).Parent).CloseInside;

end;


procedure TOM_STDCPTA.ChargeStandardCompta;
var
Q1 :TQuery;
begin

  Q1 := OpenSql ('SELECT * FROM GENERAUXREF WHERE GER_NUMPLAN='+IntToStr (NumPlanCompte), TRUE);
  if Q1.EOF then
  LoadStandardRef(NumPlanCompte,'GENERAUXREF','PLANREF');
  ferme(Q1);

  LoadStandardCompta(NumPlanCompte,'GENERAUX','GENERAUXREF');
  LoadStandardCompta(NumPlanCompte,'JOURNAL','JALREF');
  LoadStandardCompta(NumPlanCompte,'GUIDE','GUIDEREF');
  LoadStandardCompta(NumPlanCompte,'ECRGUI','ECRGUIREF');
  LoadStandardCompta(NumPlanCompte,'ANAGUI','ANAGUIREF');
  LoadStandardCompta(NumPlanCompte,'TIERS','TIERSREF');
  LoadStandardCompta(NumPlanCompte,'CORRESP','CORRESPREF');
  LoadStandardCompta(NumPlanCompte,'RUPTURE','RUPTUREREF');
  LoadStandardCompta(NumPlanCompte,'REFAUTO','REFAUTOREF');
  LoadStandardMaj(NumPlanCompte,'AXE','AXEREF','AXE',' ');
  LoadStandardCompta(NumPlanCompte,'SECTION','SECTIONREF');
  LoadStandardMaj(NumPlanCompte,'CHOIXCOD','CHOIXCODREF','TYPE;CODE', ' ');
  LoadStandardCompta(NumPlanCompte,'VENTIL','VENTILREF');
  LoadStandardCompta(NumPlanCompte,'NATCPTE','NATCPTEREF');
  LoadStandardCompta(NumPlanCompte,'FILTRES','FILTRESREF');
  LoadStandardMaj(NumPlanCompte,'LISTE','LISTEREF','LISTE',' ');
  // CA - 02/04/2002
  LoadStandardCompta(NumPlanCompte, 'MODEPAIE', 'MODEPAIEREF');
  LoadStandardCompta(NumPlanCompte, 'MODEREGL', 'MODEREGLREF');
  // Fin CA - 02/04/2002
  ChargeAxe;
end;

procedure TOM_STDCPTA.ChargeAxe;
Var
Q1,Q2   : TQuery;
Tobaxe  : TOB;
begin
  Q1 := OpenSQL('SELECT * FROM AXEREF Where XRE_NUMPLAN="'+ IntToStr (NumPlanCompte)+'"',True);
  if Q1.EOF then
  begin
    Ferme (Q1);
    Tobaxe :=TOB.Create('AXE',Nil,-1) ;
    Q2 := OpenSQLCom('SELECT * FROM AXE',True);
    Q2.First;
    while not Q2.EOF do
    begin
    Tobaxe.PutValue('X_AXE', Q2.FindField('X_AXE').Asstring);
    Tobaxe.PutValue('X_LIBELLE', Q2.FindField('X_LIBELLE').Asstring);
    Tobaxe.PutValue('X_COMPTABLE', Q2.FindField('X_COMPTABLE').Asstring);
    Tobaxe.PutValue('X_CHANTIER', Q2.FindField('X_CHANTIER').Asstring);
    Tobaxe.PutValue('X_MODEREOUVERTURE', Q2.FindField('X_MODEREOUVERTURE').Asstring);
    Tobaxe.PutValue('X_SECTIONATTENTE', Q2.FindField('X_SECTIONATTENTE').Asstring);
    Tobaxe.PutValue('X_REGLESAISIE', Q2.FindField('X_REGLESAISIE').Asstring);
    Tobaxe.PutValue('X_ABREGE', Q2.FindField('X_ABREGE').Asstring);
    Tobaxe.PutValue('X_LONGSECTION', Q2.FindField('X_LONGSECTION').Asinteger);
    Tobaxe.PutValue('X_BOURREANA', Q2.FindField('X_BOURREANA').Asstring);
    Tobaxe.PutValue('X_SOCIETE', Q2.FindField('X_SOCIETE').Asstring);
    Tobaxe.PutValue('X_STRUCTURE', Q2.FindField('X_STRUCTURE').Asstring);
    Tobaxe.PutValue('X_GENEATTENTE', Q2.FindField('X_GENEATTENTE').Asstring);
    Tobaxe.PutValue('X_CPTESTRUCT', Q2.FindField('X_CPTESTRUCT').Asstring);
    Tobaxe.PutValue('X_FERME', Q2.FindField('X_FERME').Asstring);
    Tobaxe.PutValue('X_SAISIETRANCHE', Q2.FindField('X_SAISIETRANCHE').Asstring);
    Tobaxe.InsertOrUpdateDB (TRUE);
    Q2.Next;
    end;
    FermeSQLCom (Q2);
    if Tobaxe<>Nil then Tobaxe.Free ;
  end
  else
  Ferme (Q1);
end;
function TOM_STDCPTA.ChargeStandard : boolean;
var bRetour : boolean;
begin
  if Blocage(['nrBatch','nrCloture'],True,'nrCloture') then
  begin result := false; exit ; end;
  if Transactions (ChargeStandardCompta, 1) <> oeOk then
  begin
    // Erreur durant la mise à jour depuis les standards
    PGIBox('Chargement du standard impossible. Veuillez recommencer.','Chargement d''un standard');
    Bloqueur('nrCloture',False) ;
    bRetour := false;
  end else
  begin
    SetParamSoc('SO_NUMPLANREF',NumPlanCompte);
    bRetour := True;
  end;
  Bloqueur('nrCloture',False) ;
  result := bRetour;
end;


procedure TOM_STDCPTA.SupprimeTableChoixCod(Code : string);
var QDos : TQuery;
begin
  QDos := OpenSQL('SELECT * FROM CHOIXCOD Where CC_TYPE="'+Code+'"',True);
  if not QDos.Eof then
      ExecuteSQL('Delete from CHOIXCOD Where CC_TYPE="'+Code+'"');
  Ferme (QDos);
end;

procedure TOM_STDCPTA.SupprimeTableStd (TableDos : string);
var QDos : TQuery;
begin
  QDos := OpenSQL('SELECT * FROM '+TableDos,True);
  if not QDos.Eof then
      ExecuteSQL('Delete from '+TableDos);
  Ferme (QDos);
end;

procedure TOM_STDCPTA.SupprimeTableChoixListe(LIST : string);
var QDos : TQuery;
begin
  QDos := OpenSQL('SELECT * FROM LISTE WHERE LI_LISTE="'+LIST+'"', TRUE);
  if not QDos.Eof then
      ExecuteSQL('Delete from LISTE Where LI_LISTE="'+LIST+'"');
  Ferme (QDos);
end;

procedure TOM_STDCPTA.LoadStandardRef (NumStd : integer; TableDos,TableStd : string);
var QDos, QStd : TQuery;
    PrefDos, PrefStd, Suffixe , ChampDos  : string;
    ValChamp : variant;
    i : integer;
    Val : string;
begin
  PrefDos := TableToPrefixe(TableDos);
  PrefStd := TableToPrefixe(TableStd);
  QStd := OpenSQL ('SELECT * FROM '+TableStd+' WHERE '+PrefStd+'_NUMPLAN='+IntToStr(NumStd),True);
  if not QStd.Eof then  // Si un tel standard existe
  begin
    InitMove(QCount(QStd),'Chargement Table : '+TableDos);
    QDos := OpenSQL('SELECT * FROM '+TableDos,False);
      while not QStd.Eof do
      begin
        QDos.Insert;
        InitNew (QDos);
        for i:=0 to QStd.FieldCount - 1 do
        begin
          Suffixe := ExtractSuffixe(QStd.Fields[i].FieldName);
          // Gestion des exceptions à la règle ici
          if ((PrefStd='PR') and (Suffixe='COMPTE')) then ChampDos := 'GER_GENERAL'
          else if ((PrefStd='PR') and (Suffixe='REPORTDETAIL')) then continue
          // Fin de la gestion des exceptions
          else ChampDos := PrefDos+'_'+Suffixe;
          if ChampDos='GER_BLOCNOTE' then continue;
          // calcul de la valeur du champ
          if ChampDos = 'PR_COMPTE' then
          ValChamp := BourreLaDonc(QStd.FindField('GER_GENERAL').AsString,fbGene)
          // Fin modification
          else ValChamp := QStd.FindField(PrefStd+'_'+Suffixe).AsVariant;
          QDos.FindField(ChampDos).AsVariant := ValChamp;
        end;
        QDos.Post;
        MoveCur(False);
        QStd.Next;
      end;
    Ferme (QDos);
    FiniMove;
  end;
  Ferme (QStd);
end;




Initialization
registerclasses([TOM_STDCPTA]) ;
end.
