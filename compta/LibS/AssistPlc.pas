unit AssistPlc;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  assist, HSysMenu, hmsgbox, StdCtrls, HTB97, ComCtrls, ExtCtrls, Hctrls,
  ActnList, Mask, HPanel, UTOB, HQuickRP, dbtables, ImgList, Menus,
  HEnt1,LicUtil,MajTable, PGIExec, ParamSoc,fe_main, HStatus,
  // Uses Commun
  FichComm,
  // Uses Compta
  Ent1, Devise, MulJal, MulGene, ParamLib, ContAbon, Guide, Exercice, AssistExo,
  InterSavSisco,AssistPL, Spin,galOutil;


type
  TFAssistPLC = class(TFAssist)
    TabSheet1                  : TTabSheet;
    SO_LGCPTEGEN               : TSpinEdit;
    HLabel1                    : THLabel;
    SO_LGCPTEAUX               : TSpinEdit;
    LSTMAQ: TListBox;
    BtSupp: TToolbarButton97;
    procedure FormShow(Sender: TObject);
    procedure bFinClick(Sender: TObject);
    procedure ChargeStandardCompta;
    procedure ChargeParamSocRef (NumPlan : integer; Charge : Boolean);
    function  ChargeStandard : boolean;
    procedure bAnnulerClick(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure LSTMAQDblClick(Sender: TObject);
    procedure LSTMAQClick(Sender: TObject);
    procedure BtSuppClick(Sender: TObject);
  private
         lgene, lgaux    : integer;
         procedure SupprimeTableChoixCod (Code : string);
         procedure SupprimeTableStd (TableDos : string);
         procedure ChargeAxe;
    { Déclarations privées }
  public
    { Déclarations publiques }
  end;

var
  FAssistPLC: TFAssistPLC;
function LanceAssistantPlc (bInit : boolean) : boolean;
procedure LoadStandardRef (NumStd : integer; TableDos,TableStd : string);

implementation
uses
TOMStdcpta,CLgCpte;

{$R *.DFM}
function LanceAssistantPlc (bInit : boolean) : boolean;
var
      Q1      : TQuery;
begin

    FAssistPLC := TFAssistPLC.Create (Application);
    try
          Q1 := OpenSQl ('SELECT * From SOCIETE', TRUE);
          if Q1.EOF then InitSocietePCL ;
          Ferme (Q1);
          Q1 := OpenSQl ('SELECT * From ETABLISS', TRUE);
          if Q1.EOF then InitEtablissement;
          Ferme (Q1);
          FAssistPLC.ShowModal;
    finally
      FAssistPLC.Free;
    end;
    if FAssistPLC.ModalResult = mrCancel then
       result := FALSE
    else
       result := TRUE;
end;

procedure TFAssistPLC.FormShow(Sender: TObject);
var
St         : string;
Q1         : Tquery;
i,Index    : integer;
begin
  inherited;
  NumPlanCompte := GetParamSoc('SO_NUMPLANREF');
  ChargeParamSocRef(NumPlanCompte, TRUE);

  if GetParamSoc ('SO_LGCPTEGEN')= 0 then
     SO_LGCPTEGEN.Value := 6
  else
     SO_LGCPTEGEN.Value := GetParamSoc ('SO_LGCPTEGEN');
  if GetParamSoc('SO_LGCPTEAUX')= 0 then
     SO_LGCPTEAUX.Value := 6
  else
     SO_LGCPTEAUX.Value := GetParamSoc ('SO_LGCPTEAUX');

  if GetParamSoc ('SO_BOURREGEN') <> '0' then
     SetParamSoc ('SO_BOURREGEN', '0');
  if GetParamSoc ('SO_BOURREAUX') <> '0' then
     SetParamSoc ('SO_BOURREAUX', '0');

      // Chargement VH
  ChargeSocieteHalley;

  Q1 := OpenSql ('SELECT * FROM STDCPTA', TRUE);
  i := 0;
  While not Q1.EOF do
  begin
       St := (Format('%.03d',[Q1.FindField('STC_NUMPLAN').AsInteger])) +
       '   ' + Q1.FindField('STC_LIBELLE').AsString;
       if  NumPlanCompte = Q1.FindField('STC_NUMPLAN').AsInteger then
       Index := i;
       LSTMAQ.Items.Add(St); inc(i);
       Q1.next;
  end;
  Ferme (Q1);

  LSTMAQ.ItemIndex := Index;


end;

procedure TFAssistPLC.bFinClick(Sender: TObject);
var
ValChamp,where        : string;
Q1                    : Tquery;
Texte                 : string;
begin
  inherited;

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

    SetParamSoc ('SO_NUMPLANREF', NumPlanCompte);
    SetParamSoc ('SO_BOURREGEN', '0');
    SetParamSoc ('SO_BOURREAUX', '0');

   ChargeParamSocRef(NumPlanCompte, FALSE);

  if GetParamSoc ('SO_BOURREGEN') <> '0' then
     SetParamSoc ('SO_BOURREGEN', '0');
  if GetParamSoc ('SO_BOURREAUX') <> '0' then
     SetParamSoc ('SO_BOURREAUX', '0');

  if  SO_LGCPTEGEN.Value <> GetParamSoc ('SO_LGCPTEGEN') then
     SetParamSoc ('SO_LGCPTEGEN', SO_LGCPTEGEN.Value);
  if  SO_LGCPTEAUX.Value <> GetParamSoc ('SO_LGCPTEAUX') then
     SetParamSoc ('SO_LGCPTEAUX', SO_LGCPTEAUX.Value);
  if (NumPlanCompte > 0 ) and (NumPlanCompte <> GetParamSoc('SO_NUMPLANREF'))then
     SetParamSoc ('SO_NUMPLANREF', NumPlanCompte);
  ChargeMagHalley;
  ChargeStandard;

  if SO_LGCPTEGEN.Value <> lgene then
     MoulinetteChangeLg('0', SO_LGCPTEGEN.Value);

     Q1 := OpenSql ('SELECT EG_GENERAL, EG_AUXILIAIRE,EG_TYPE, EG_GUIDE, EG_NUMLIGNE from ECRGUI', TRUE);
     while not Q1.EOF do
     begin
       ValChamp := BourreLaDonc(Q1.Fields[0].asstring,fbGene);
       if Length (ValChamp) > GetParamSoc ('SO_LGCPTEGEN') then
       begin
       ValChamp := copy(Q1.Fields[0].asstring,0,SO_LGCPTEGEN.value);
       where := '" where EG_TYPE="'+ Q1.findfield('EG_Type').asstring +'" and EG_GUIDE="'+
       Q1.findfield ('EG_GUIDE').asstring + '" and EG_NUMLIGNE=' + IntToStr(Q1.findfield ('EG_NUMLIGNE').asinteger);
       ExecuteSQL('UPDATE  ECRGUI SET EG_GENERAL="'+ ValChamp+
       where) ;
       end;
       ValChamp := BourreLaDonc(Q1.Fields[1].asstring,fbGene);
       if Length (ValChamp) > GetParamSoc ('SO_LGCPTEAUX') then
       begin
       ValChamp := copy(Q1.Fields[0].asstring,0,SO_LGCPTEAUX.value);
       where := '" where EG_TYPE="'+ Q1.findfield('EG_Type').asstring +'" and EG_GUIDE="'+
       Q1.findfield ('EG_GUIDE').asstring + '" and EG_NUMLIGNE=' + IntToStr(Q1.findfield ('EG_NUMLIGNE').asinteger);
       ExecuteSQL('UPDATE  ECRGUI SET EG_AUXILIAIRE="'+ ValChamp+
       where) ;
       end;
       Q1.next;
     end;
     Ferme(Q1);

   if SO_LGCPTEAUX.Value <> lgaux  then
     MoulinetteChangeLgAux('0', SO_LGCPTEAUX.Value);
   ModalResult := MROK;

end;

procedure TFAssistPLC.ChargeParamSocRef (NumPlan : integer; Charge : Boolean);
var QParamSocRef : TQuery ;
    StVal : string;
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
    SO_LGCPTEGEN.value :=  StrToInt(QParamSocRef.FindField('PRR_SOCDATA').AsString);
    lgene := SO_LGCPTEGEN.value;
    SetParamSoc ('SO_LGCPTEGEN', SO_LGCPTEGEN.value);
    end;
    if QParamSocRef.FindField('PRR_SOCNOM').AsString = 'SO_LGCPTEAUX' then
    begin
    if Charge then
    SO_LGCPTEAUX.value :=  StrToInt(QParamSocRef.FindField('PRR_SOCDATA').AsString);
    lgaux := SO_LGCPTEAUX.value;
    SetParamSoc ('SO_LGCPTEAUX', SO_LGCPTEAUX.value);
    end;
    SetParamSoc ( QParamSocRef.FindField('PRR_SOCNOM').AsString, StVal);
    QParamSocRef.Next ;
    MoveCur (False);
  end;
  Ferme(QParamSocRef);
  FiniMove;
end;

procedure TFAssistPLC.ChargeStandardCompta;
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
//  LoadStandardCompta(NumPlanCompte,'AXE','AXEREF');
  LoadStandardMaj(NumPlanCompte,'AXE','AXEREF','AXE',' ');
  LoadStandardCompta(NumPlanCompte,'SECTION','SECTIONREF');
//  LoadStandardCompta(NumPlanCompte,'CHOIXCOD','CHOIXCODREF');
  LoadStandardMaj(NumPlanCompte,'CHOIXCOD','CHOIXCODREF','TYPE;CODE', ' ');
  LoadStandardCompta(NumPlanCompte,'VENTIL','VENTILREF');
  LoadStandardCompta(NumPlanCompte,'NATCPTE','NATCPTEREF');
  LoadStandardCompta(NumPlanCompte,'FILTRES','FILTRESREF');
//  LoadStandardCompta(NumPlanCompte,'LISTE','LISTEREF');
  LoadStandardMaj(NumPlanCompte,'LISTE','LISTEREF','LISTE',' ');
  ChargeAxe;
end;

procedure TFAssistPLC.ChargeAxe;
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
function TFAssistPLC.ChargeStandard : boolean;
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


procedure TFAssistPLC.SupprimeTableChoixCod(Code : string);
var QDos : TQuery;
begin
  QDos := OpenSQL('SELECT * FROM CHOIXCOD Where CC_TYPE="'+Code+'"',True);
  if not QDos.Eof then
      ExecuteSQL('Delete from CHOIXCOD Where CC_TYPE="'+Code+'"');
  Ferme (QDos);
end;

procedure TFAssistPLC.SupprimeTableStd (TableDos : string);
var QDos : TQuery;
begin
  QDos := OpenSQL('SELECT * FROM '+TableDos,True);
  if not QDos.Eof then
      ExecuteSQL('Delete from '+TableDos);
  Ferme (QDos);
end;

procedure LoadStandardRef (NumStd : integer; TableDos,TableStd : string);
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



procedure TFAssistPLC.bAnnulerClick(Sender: TObject);
begin
  inherited;
ModalResult := mrCancel;
end;

procedure TFAssistPLC.FormKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  inherited;
if Key=VK_F5 then AGLLanceFiche('CP','STDPLCPTE','','','');
end;


procedure TFAssistPLC.LSTMAQDblClick(Sender: TObject);
var
Text, Lib       : string;
Index, Num      : integer;
begin
  inherited;
     Index := LSTMAQ.ItemIndex;
     if Index >= 0 then
     Text := LSTMAQ.Items[Index];
     Text := copy (Text, 1, 3);
     Num := StrToInt (Text);
     Lib := AGLLanceFiche('CP','STDPLCPTE',IntToStr(Num), '','');
     LSTMAQ.Items[Index] := Text + '   ' + Lib;
end;

procedure TFAssistPLC.LSTMAQClick(Sender: TObject);
var
Text            : string;
Index, Num      : integer;
begin
  inherited;
     Index := LSTMAQ.ItemIndex;
     if Index >= 0 then
     Text := LSTMAQ.Items[Index];
     Text := copy (Text, 1, 3);
     NumPlanCompte := StrToInt (Text);
end;

procedure TFAssistPLC.BtSuppClick(Sender: TObject);
var
Predef            : string;
Where             : string;
begin
  inherited;
if not IsSuperviseur(TRUE) then
begin
PGIInfo('Vous n''êtes pas un utilisateur privilégié','') ; exit;
end;
if NumPlanCompte < 21 then Predef := 'CEG'
else Predef := 'STD';
if not FileExists('STDCEGID.DAT') then
if NumPlanCompte <= 20 then
begin PGIInfo('Plan type CEGID,impossible de supprimer','') ;  exit;end;

Where := 'Where STC_NUMPLAN='+IntToStr (NumPlanCompte)+ ' and STC_PREDEFINI="'+Predef+'"';
ExecuteSQL('DELETE FROM STDCPTA '+Where);

SupprimeDonneStd ('PARSOCREF',NumPlanCompte);
SupprimeDonneStd ('GENERAUXREF',NumPlanCompte);
SupprimeDonneStd ('JALREF',NumPlanCompte);
SupprimeDonneStd ('GUIDEREF',NumPlanCompte);
SupprimeDonneStd ('ECRGUIREF',NumPlanCompte);
SupprimeDonneStd ('ANAGUIREF',NumPlanCompte);
SupprimeDonneStd ('TIERSREF',NumPlanCompte);
SupprimeDonneStd ('CORRESPREF',NumPlanCompte);
SupprimeDonneStd ('RUPTUREREF',NumPlanCompte);
SupprimeDonneStd ('REFAUTOREF',NumPlanCompte);
SupprimeDonneStd ('AXEREF',NumPlanCompte);
SupprimeDonneStd ('SECTIONREF',NumPlanCompte);
SupprimeDonneStd ('CHOIXCODREF',NumPlanCompte);
SupprimeDonneStd ('VENTILREF',NumPlanCompte);
SupprimeDonneStd ('NATCPTEREF',NumPlanCompte);
SupprimeDonneStd ('FILTRESREF',NumPlanCompte);
SupprimeDonneStd ('LISTEREF',NumPlanCompte);

LSTMAQ.Items.Delete(LSTMAQ.ItemIndex);
end;


end.
