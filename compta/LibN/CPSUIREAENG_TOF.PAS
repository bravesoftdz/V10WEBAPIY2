{***********UNITE*************************************************
Auteur  ...... : BELLARD Frédéric
Créé le ...... : 27/06/2006
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : CPSUIREAENG_QR1 ()
Mots clefs ... : TOF;CPSUIREAENG_QR1
*****************************************************************}
Unit CPSUIREAENG_TOF ;

Interface

Uses StdCtrls,
     Controls,
     Classes,                             
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,
     Fe_Main,     // AGLLanceFiche
{$else}
     eMul,
     uTob,
     MaineAGL,
{$ENDIF}
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     UTOF,
     TofMeth,
     Ent1,
     ParamSoc,
     ULibWindows;

procedure CPLanceFiche_CPSuiReaEng;

Type
  TOF_CPSUIREAENG_QR1 = Class (TOF_Meth)
    private
      ChoixAxe : THValComboBox;
      AxeGT : THEdit;
      Compte1 : THEdit;
      Compte2 : THEdit;
      Etabl : THValComBoBox;
      Journaux : THMultiValComboBox;
      LJourn : THEdit;
      Statut : THMultiValComboBox;
      LStat : THEdit;
      DatePrev1 : THEdit;
      DatePrev2 : THEdit;
      DateEcr1 : THEdit;
      DateEcr2 : THEdit;
      RefCom1 : THEdit;
      RefCom2 : THEdit;
      DateEch1 : THEdit;
      DateEch2 : THEdit;
      ConditionSQLEcrS : THEdit;
      ConditionSQLEcrF : THEdit;
      ConditionSQLEcrW : THEdit;
      ConditionSQLEcrO : THEdit;

      function TableCompte : string;
      function ChampCompte : string;
      function ChampEcriture : string;
      function ChampLibelle  : string;
      function  GetMinMaxCompte(stTable, stCol, stDefaut : String) : String;
      procedure ChoixAxeOnChange(Sender: TObject);
      procedure GenererConditionSQL;
    public
      procedure OnLoad                   ; override ;
      procedure OnArgument (S : String ) ; override ;
  end ;

Implementation

procedure CPLanceFiche_CPSuiReaEng;
begin
  // Etat non utilisable sous ORACLE7
  if V_PGI.Driver = dbORACLE7 then
    PGIInfo( 'Cet état n''est pas utilisable sous Oracle 7. Veuillez utiliser l''état correspondant dans le menu Autres éditions' , 'Justificatif de soldes' )
  else
    AGLLanceFiche('CP', 'CPSUIREAENG_QR1', '', '', '') ;
end;

function TOF_CPSUIREAENG_QR1.TableCompte : string;
begin
  Case ChoixAxe.ItemIndex of
    0..1 : Result := 'TIERS';
    2 : Result := 'GENERAUX';
    end;
end;

function TOF_CPSUIREAENG_QR1.ChampEcriture : string;
var
  Champ : string;
begin
  Case ChoixAxe.ItemIndex of
    0..1 : Champ := ' E_AUXILIAIRE';
    2    : Champ := ' E_GENERAL';
    end;
  Result := Champ;
end;

function TOF_CPSUIREAENG_QR1.ChampCompte : string;
var
  Champ : string;
begin
  Case ChoixAxe.ItemIndex of
    0..1 : Champ := ' T_AUXILIAIRE';
    2    : Champ := ' G_GENERAL';
    end;
  Result := Champ;
end;

function TOF_CPSUIREAENG_QR1.GetMinMaxCompte(stTable, stCol, stDefaut : String) : String;
var
  Q : TQuery;
  stWhere : String;
begin
  StWhere := '';
  if stTable = 'TIERS' then begin
    if ChoixAxe.itemindex=0 then
      stWhere := ' WHERE T_NATUREAUXI="CLI" '
    else
      stWhere := ' WHERE T_NATUREAUXI="FOU" ';
    end;

  Q := OpenSQL('SELECT ' + stCol +' CODE FROM ' + stTable + stWhere,True);
  if not Q.Eof then
    Result := Q.FindField('CODE').asString
  else
    Result := stDefaut;
  Ferme(Q);
end;

procedure TOF_CPSUIREAENG_QR1.ChoixAxeOnChange(Sender: TObject);
begin
  if ChoixAxe.ItemIndex=0 then
    Compte1.DataType := 'TZTFOURN'
  else if ChoixAxe.ItemIndex=1 then
    Compte1.DataType := 'TZTCLIENT'
  else if ChoixAxe.itemindex=2 then
    Compte1.DataType := 'TZGENERAL';

  Compte2.DataType:=Compte1.DataType;
  Compte1.Text:='';
  Compte2.Text:='';

  SetControlText('AXEGT',ChoixAxe.Items[ChoixAxe.itemindex]);
end;

procedure TOF_CPSUIREAENG_QR1.GenererConditionSQL;
var
  StSelect, StFrom, StWhere, StOrder : string;
  stWhereTraiteJ, stWhereTraiteS : string ;
  StJournal : string;
  StStatut : string;
begin
  Inherited ;
  { Construction de la clause Select de la SQL }
  StSelect := '';
  StFrom := '';
  StWhere := '';
  StOrder := '';

  StSelect := ', ' + ChampEcriture + ' RUPCOMPTE , ' + ChampLibelle + ' RUPLIBELLE ' ;

  { Construction de la clause Where de la SQL }
  //Condition sur le plan de comptes
  if ChoixAxe.itemindex<2
    then StFrom := ' INNER JOIN TIERS ON E_AUXILIAIRE=T_AUXILIAIRE '
    else StFrom := ' LEFT JOIN GENERAUX ON E_GENERAL=G_GENERAL ' ;

  if ChoixAxe.itemindex<2 then
    stWhere := ' E_AUXILIAIRE<>"" '
  else
    stWhere := ' E_AUXILIAIRE="" ';

  if ChoixAxe.itemindex=0 then
    stWhere := StWhere + ' AND T_NATUREAUXI="FOU" ';

  if ChoixAxe.itemindex=1 then
    stWhere := StWhere + ' AND T_NATUREAUXI="CLI" ';

  // Condition sur les comptes
    stWhere := StWhere + ' AND ' + ChampEcriture + '>="' + GetControlText('COMPTE') +
  					 '" AND' + ChampEcriture + '<="' + GetControlText('COMPTE_') + '"';

  // Condition sur l'etablissement
  if Etabl.itemIndex > 0 then
    stWhere := stWhere + ' AND E_ETABLISSEMENT = "' + Etabl.Value + '"';

  // Condition sur le journal
  TraductionTHMultiValComboBox( Journaux, stWhereTraiteJ , stJournal, 'E_JOURNAL', false );
  stWhere := stWhere + ' AND ' + stWhereTraiteJ;
  if LJOURN<>nil then begin
    if stJournal='<<Tous>>' then
      SetControlText('LJOURN',TraduireMemoire(stJournal))
    else
      SetControlText('LJOURN',Journaux.value);
    end;

  // Condition sur le statut
  if Statut<>nil then begin
    stStatut := Statut.Value;

    if LSTAT<>nil then begin
      if stStatut='' then
        SetControlText('LSTAT',TraduireMemoire('<<Tous>>'))
      else
        SetControlText('LSTAT',Statut.value);
      end;

    if stStatut <> '' then begin
      stWhereTraiteS := '';
      while stStatut <> '' do
        if stWhereTraiteS = '' then
          stWhereTraiteS := ' AND CEN_STATUTENG IN ("' + ReadTokenSt(stStatut) + '"'
        else
          stWhereTraiteS := stWhereTraiteS + ',"' + ReadTokenSt(stStatut) + '"';
      stWhere := stWhere + stWhereTraiteS + ')';
      end;
    end;

  //Condition sur la date d'enregistrement
  StWhere := StWhere + ' And (E_DATECOMPTABLE>="'+USDateTime(StrToDate(DateEcr1.Text))+'" ' +
                     ' And E_DATECOMPTABLE<="'+USDateTime(StrToDate(DateEcr2.Text))+'") ';

  
  //Condition sur la réf commande
  if GetControlText('REFCOM')<>'' then
    StWhere := StWhere + ' And (E_REFINTERNE>="' + GetControlText('REFCOM') + '"' +
                         ' And E_REFINTERNE<="' + GetControlText('REFCOM_') + '") ';

  //Condition sur la date prévisionnelle
  StWhere := StWhere + ' And (E_DATEREFEXTERNE>="'+USDateTime(StrToDate(DatePrev1.Text))+'" ' +
                     ' And E_DATEREFEXTERNE<="'+USDateTime(StrToDate(DatePrev2.Text))+'") ';

  //Condition sur la date d'échéance prévisionnelle
  StWhere := StWhere + ' And (E_DATEECHEANCE>="'+USDateTime(StrToDate(DateEch1.Text))+'" ' +
                     ' And E_DATEECHEANCE<="'+USDateTime(StrToDate(DateEch2.Text))+'") ';

  { Construction de la clause Order By de la SQL }
  StOrder := StOrder + ' Order By ' + ChampEcriture + ', E_REFINTERNE, E_NUMEROPIECE ';

  ConditionSQLEcrS.Text := StSelect;
  ConditionSQLEcrF.Text := StFrom;
  ConditionSQLEcrW.Text := StWhere;
  ConditionSQLEcrO.Text := StOrder;
end;

procedure TOF_CPSUIREAENG_QR1.OnLoad ;
begin
  Inherited ;
  if (GetControlText('COMPTE')='') then
    SetControlText('COMPTE',GetMinMaxCompte(TableCompte, 'MIN(' + ChampCompte + ')','0'));
  if (GetControlText('COMPTE_')='') then
    SetControlText('COMPTE_',GetMinMaxCompte(TableCompte, 'MAX(' + ChampCompte + ')','ZZZZZZZZZZZZZZZZZ'));

  // Remplir les conditions SQL
  GenererConditionSQL;
end;

procedure TOF_CPSUIREAENG_QR1.OnArgument (S : String ) ;
begin
  Inherited ;

  ChoixAxe := THValComboBox(GetControl('CHOIXAXE'));
  AxeGT := THEdit(GetControl('AXEGT'));
  Compte1 := THEdit(GetControl('COMPTE'));
  Compte2 := THEdit(GetControl('COMPTE_'));
  Etabl := THValComBoBox(GetControl('ETABL'));
  Journaux := THMultiValComboBox(GetControl('JOURNAL'));
  LJourn := THEdit(GetControl('LJOURN'));
  Statut := THMultiValComboBox(GetControl('STATUT'));
  LStat := THEdit(GetControl('LSTAT'));
  DatePrev1 := THEdit(GetControl('DATEPREV'));
  DatePrev2 := THEdit(GetControl('DATEPREV_'));
  DateEcr1 := THEdit(GetControl('DATEENR'));
  DateEcr2 := THEdit(GetControl('DATEENR_'));
  RefCom1 := THEdit(GetControl('REFCOM'));
  RefCom2 := THEdit(GetControl('REFCOM_'));
  DateEch1 := THEdit(GetControl('DATEECH'));
  DateEch2 := THEdit(GetControl('DATEECH_'));
  ConditionSQLEcrS := THEdit(GetControl('SQLECRS'));
  ConditionSQLEcrF := THEdit(GetControl('SQLECRF'));
  ConditionSQLEcrW := THEdit(GetControl('SQLECRW'));
  ConditionSQLEcrO := THEdit(GetControl('SQLECRO'));

  if ChoixAxe<>nil then begin
    ChoixAxe.OnChange := ChoixAxeOnChange;
    ChoixAxe.itemindex:=0;
    ChoixAxe.OnChange(nil);
    end;
  if Compte1<>nil then
    Compte1.DataType:='TZTFOURN';
  if Compte2<>nil then
    Compte2.DataType:='TZTFOURN';

  if Etabl<>nil then
    Etabl.itemindex:=0;
  PositionneEtabUser(Etabl, False);

  if Journaux<>nil then
    Journaux.Value:=TraduireMemoire('<<Tous>>');

  if Statut<>nil then
    Statut.Value:=TraduireMemoire('<<Tous>>');

  if DatePrev1<>nil then
    DatePrev1.Text:=DateToStr( iDate1900 );

  if DatePrev2<>nil then
    DatePrev2.Text:=DateToStr( iDate2099 );

  if DateEcr1<>nil then
    DateEcr1.Text:=DateToStr( iDate1900 );

  if DateEcr2<>nil then
    DateEcr2.Text:=DateToStr(V_PGI.DateEntree);

  if DateEch1<>nil then
    DateEch1.Text:=DateToStr( iDate1900 );

  if DateEch2<>nil then
    DateEch2.Text:=DateToStr( iDate2099 );
end ;

function TOF_CPSUIREAENG_QR1.ChampLibelle: string;
begin
  Case ChoixAxe.ItemIndex of
    2 :   result := 'G_LIBELLE' ;
    else  result := 'T_LIBELLE' ;
  end ;
end;

Initialization
  registerclasses ( [ TOF_CPSUIREAENG_QR1 ] ) ;
end.
