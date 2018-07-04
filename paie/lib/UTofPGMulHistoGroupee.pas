{***********UNITE*************************************************
Auteur  ...... : JL
Créé le ...... : 25/05/2005
Modifié le ... :   /  /
Description .. : Saisie historique
Mots clefs ... : PAIE
*****************************************************************
PT1 15/11/2005 JL V_650  ajout fonction rend libelle pour mettre a jour libellé des champs libres.
PT2 11/10/2006 JL V_700 FQ 13556 Ne pas afficher écran de saisie si pas de sélection
PT3 15/11/2006 JL V_70 FQ 13397 Fetchlestous déplacé
PT4 02/08/2007 JL V_80 FQ 14622 Initialisation de la date avec début de mois
PT5 21/12/2007 GGU V_80 FQ 14684 Proposer une Saisie groupée pour les salariés dont les critères de population sont non renseignés.
PT6 09/01/2008 GGU V_81 FQ 13505 "date libre" à rajouter dans la MAJ FICHE SALARIE
}

unit UTofPGMulHistoGroupee;

interface
uses  StdCtrls,Controls,Classes,sysutils,
{$IFNDEF EAGLCLIENT}
      db,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}HDB,Mul,Fe_Main,
{$ELSE}
      eMul,MainEAGL,
{$ENDIF}
      HCtrls,HEnt1,HMsgBox,UTOF,UTOB,HTB97,
      HStatus,ed_tools,P5Def,HQry,PgOutilsHistorique,
      EntPaie,PgOutils2;

Type
     TOF_PGMULHISTOGROUPEE = Class (TOF)
       procedure OnArgument (stArgument: String);        Override;
       procedure ExitEdit(Sender: TObject);
       procedure OnLoad ;                               Override;
       procedure OnActivate(Sender : TObject); //PT5
   private
       Argument,StTemp : string;
       AutoValide : Boolean;
       {$IFNDEF EAGLCLIENT}
       Liste : THDBGrid;
       {$ELSE}
       Liste : THGrid;
       {$ENDIF}
       procedure OnClickSalarieSortie(Sender: TObject);
       procedure GrilleDblClick(Sender : TObject);
       procedure ConstruireTobHistorique;
       procedure ConstruireTobHistoriqueModif;
       procedure RemplirComboChoix;
       Procedure RendLibelle(Num : Integer;var Libelle : String);
       Procedure NewHisto(Sender : TObject);
    END ;
var PGTobHistoSal,PGTobLesChamps : Tob;
    PGDateDebHisto,PGDateFinHisto : TDateTime;

implementation

uses
  ParamSoc, //PT6
  StrUtils;

procedure TOF_PGMULHISTOGROUPEE.GrilleDblClick(Sender : TObject);
var MultiC : THMultivALCombobox;
begin
     MultiC := THMultiValCombobox(GetControl('LISTECHAMP'));
     If MultiC.TEXT = '' then
     begin
          PGIBox('Vous devez choisir au moins une information à modifier',Ecran.Caption);
          Exit;
     end;
     if (Liste.NbSelected = 0) and (TFMul(Ecran).BSelectAll.Down = False) then       //PT2
     begin
          PGIBox('Aucun élément sélectionné', Ecran.Caption);
          Exit;
     end;
     PGTobHistoSal := tob.Create('SaisieHisto',Nil,-1);
//     ConstruireTobHistoriqueModif;
//     AGLLanceFiche('PAY','HISTOGROUPEEMODIF','','',GetControlText('LISTECHAMP')+';'+'MODIFICATION');
     ConstruireTobHistorique;
     AGLLanceFiche('PAY','HISTOGROUPEE','','',GetControlText('LISTECHAMP')+';'+'CREATION');
     Liste.ClearSelected;
     PGTobHistoSal.Free;
     PGTobLesChamps.Free;
end;

procedure TOF_PGMULHISTOGROUPEE.OnArgument (stArgument: String);
var
Defaut: THEdit;
Check : TCheckBox;
Num : Integer;
Q : Tquery;
BNew : TToolBarButton97;
Arg, tempArg : String; //PT5
i : Integer; //PT6
begin
//Debut PT6
  { Mise à jour de PAIEPARIM avec les dates libres }
  ExecuteSQL('DELETE FROM PAIEPARIM WHERE PAI_IDENT BETWEEN 2000 AND 2019');
  for i := 1 to GetParamSocSecur('SO_PGNBDATE', 0) do
  begin
    ExecuteSQL('Insert into PAIEPARIM (PAI_PREDEFINI, PAI_NODOSSIER, PAI_IDENT, PAI_COLONNE, PAI_PREFIX, '
      +'PAI_SUFFIX, PAI_LIENASSOC, PAI_LETYPE, PAI_CHOIX, PAI_LIBELLE, PAI_UTILISABLEPOP, '
      +'PAI_CRITERETABLE, PAI_PGTYPEUTILIS, PAI_HISTORIQUE) '
      +'Values ("CEG", "000000", '+IntToStr(2000+i)+', "DATELIBRE'+IntToStr(i)+'", "PSA", '
      +'"DATELIBRE'+IntToStr(i)+'", "", "D", "X", "'+GetParamSocSecur('SO_PGLIBDATE'+IntToStr(i), '')+'", "", '
      +'"-", "", "-")');
  end;
  for i := 1 to GetParamSocSecur('SO_PGNBCOCHE', 0) do
  begin
    ExecuteSQL('Insert into PAIEPARIM (PAI_PREDEFINI, PAI_NODOSSIER, PAI_IDENT, PAI_COLONNE, PAI_PREFIX, '
      +'PAI_SUFFIX, PAI_LIENASSOC, PAI_LETYPE, PAI_CHOIX, PAI_LIBELLE, PAI_UTILISABLEPOP, '
      +'PAI_CRITERETABLE, PAI_PGTYPEUTILIS, PAI_HISTORIQUE) '
      +'Values ("CEG", "000000", '+IntToStr(2010+i)+', "BOOLLIBRE'+IntToStr(i)+'", "PSA", '
      +'"BOOLLIBRE'+IntToStr(i)+'", "", "B", "X", "'+GetParamSocSecur('SO_PGLIBCOCHE'+IntToStr(i), '')+'", "", '
      +'"-", "", "-")');
  end;
  AvertirTable('PGCHAMPDISPO');
  SetControlProperty('LISTECHAMP', 'DataType', 'PGCHAMPDISPO');
//Fin PT6
Inherited ;
//début PT5
AutoValide := False;
if stArgument <> '' then
begin
  Arg := READTOKENPipe(stArgument, '|');
  if pos('CRITERE=', Arg) > 0 then
  begin
    Arg := RightStr(Arg, Length(Arg) - 8);
    SetControlText('LISTECHAMP',Arg);
    AutoValide := True;
  end;
  Arg := READTOKENPipe(stArgument, '|');
  if pos('SALARIES=', Arg) > 0 then
  begin
    Arg := RightStr(Arg, Length(Arg) - 9);
    tempArg := READTOKENPipe(Arg, ';');
    while tempArg <> '' do
    begin
      StTemp := StTemp + ','+tempArg;
      tempArg := READTOKENPipe(Arg, ';');
    end;
    StTemp := copy(StTemp, 2, length(StTemp));
    StTemp := 'AND PSA_SALARIE IN (' + StTemp + ') ';
  end;
end;
TFMul(Ecran).OnActivate := OnActivate;  
//Fin PT5
//PGSalariesHisto;
RemplirComboChoix;
//Supp date pour saisie histo date a date
//SetControlVisible('BInsert',True);
//BNew := TToolBarButton97(GetControl('BInsert'));
//If BNew <> Nil then BNew.OnClick := NewHisto;
TfMul(Ecran).Caption := 'Mise à jour salarié';
UpdateCaption(TFMul(Ecran));
{$IFNDEF EAGLCLIENT}
Liste := THDBGrid(GetControl('FListe'));
{$ELSE}
Liste := THGrid(GetControl('FListe'));
{$ENDIF}
If Liste <> Nil Then Liste.OnDblClick := GrilleDblClick ;
 SetControlvisible('DATEARRET',True);
 SetControlvisible('TDATEARRET',True);
 SetControlEnabled('DATEARRET',False);
 SetControlEnabled('TDATEARRET',False);
 Check:=TCheckBox(GetControl('CKSORTIE'));
 if Check=nil then
   Begin
    SetControlVisible('DATEARRET',False);
    SetControlVisible('TDATEARRET',False);
    End
  else
    Check.OnClick:=OnClickSalarieSortie;
 For Num := 1 to VH_Paie.PGNbreStatOrg do
 begin
 if Num >4 then Break;
 VisibiliteChampSalarie (IntToStr(Num),GetControl ('PSA_TRAVAILN'+IntToStr(Num)),GetControl ('TPSA_TRAVAILN'+IntToStr(Num)));
 end;
 VisibiliteStat (GetControl ('PSA_CODESTAT'),GetControl ('TPSA_CODESTAT')) ; //PT3
 Defaut:=ThEdit(getcontrol('PSA_SALARIE'));
 If Defaut<>nil then Defaut.OnExit:=ExitEdit;
end;

procedure TOF_PGMULHISTOGROUPEE.ExitEdit(Sender: TObject);
var edit : thedit;
begin
edit:=THEdit(Sender);
if edit <> nil then	//AffectDefautCode que si gestion du code salarié en Numérique
    if (VH_Paie.PgTypeNumSal='NUM') and (length(Edit.text)<11) and (isnumeric(edit.text)) then
    edit.text:=AffectDefautCode(edit,10);
end; 


procedure TOF_PGMULHISTOGROUPEE.OnLoad ;
var
Action : THEdit;
sal : thedit;
DateArret : TdateTime;
StDateArret : string;
begin
if  TCheckBox(GetControl('CKSORTIE'))<>nil then
  Begin
  {PT6 mise en commentaire}
  if (GetControlText('CKSORTIE')='X') and (IsValidDate(GetControlText('DATEARRET')))then   //DEB PT2
     Begin
     DateArret:=StrtoDate(GetControlText('DATEARRET'));
     StDateArret:=' AND (PSA_DATESORTIE>="'+UsDateTime(DateArret)+'" OR PSA_DATESORTIE="'+UsdateTime(Idate1900)+'" OR PSA_DATESORTIE IS NULL) ';
     // PT5 13/05/2002 PH V575 Ajout controle date entrée par rapport à la date d'arreté
     StDateArret:=StDateArret + ' AND PSA_DATEENTREE <="'+UsDateTime(DateArret)+'"';
     SetControlText('XX_WHERE',StTemp+StDateArret);
     End                                             //FIN PT2
  else
     SetControlText('XX_WHERE',StTemp);
  End
Else
  StDateArret:='';
end;

procedure TOF_PGMULHISTOGROUPEE.OnClickSalarieSortie(Sender: TObject);
begin
//DEB PT7
SetControlenabled('DATEARRET',(GetControltext('CKSORTIE')='X'));
SetControlenabled('TDATEARRET',(GetControltext('CKSORTIE')='X'));
//FIN PT7
end;

procedure TOF_PGMULHISTOGROUPEE.ConstruireTobHistorique;
var Q_Mul: THQuery;
    Salarie : String;
    Q : TQuery;
    LeChamp,ChampsSelect,IDent,Colonne : String;
    DateDebut,DateFin : TDateTime;
    TH,TLC,TobHisto : Tob;
    c,h,i : integer;
    DHisto : TDateTime;
    IdH,IdS,LeType,Prefixe,Suffixe ,Tablette,Libelle : String;
    NumChamp : Integer;
    MultiC : THMultiValCombobox;
    ExisteHisto : Boolean;
    DateApplic : TDateTime;
begin
  DateApplic := DebutDeMois(V_PGI.DateEntree);//PT4
  MultiC := THMultiValCombobox(GetControl('LISTECHAMP'));
  NumChamp := 1;
  DateDebut := StrToDate(GetControlText('DATEDEBUT'));
  DateFin := StrToDate(GetControlText('DATEFIN'));
  PGDateDebHisto := DateDebut;
  PGDateFinHisto := DateFin;
  ChampsSelect := GetControlText('LISTECHAMP');
  PGTobLesChamps := Tob.Create('LesChamps',Nil,-1);
  LeChamp := ReadTokenpipe(ChampsSelect,';');
  If not MultiC.Tous then
  begin
       PGTobLesChamps.LoadDetailDB('LesChamps','','',Nil,False);
       While LeChamp <> '' do
       begin
            TLC := Tob.Create('UnChamp',PGTobLesChamps,-1);
            Q := OpenSQL('SELECT PAI_PREFIX,PAI_SUFFIX,PAI_IDENT,PAI_LETYPE,PAI_LIBELLE,PAI_LIENASSOC,PAI_COLONNE FROM PAIEPARIM WHERE PAI_IDENT="'+LeChamp+'"',True);
            If Not Q.Eof then
            begin
                 Prefixe := Q.FindField('PAI_PREFIX').AsString;
                 Suffixe := Q.FindField('PAI_SUFFIX').AsString;
                 Ident := Q.FindField('PAI_IDENT').AsString;
                 LeType := Q.FindField('PAI_LETYPE').AsString;
                 Tablette := Q.FindField('PAI_LIENASSOC').AsString;
                 Libelle := Q.FindField('PAI_LIBELLE').AsString;
                 Colonne := Q.FindField('PAI_COLONNE').AsString;
            end;
            Ferme(Q);
            RendLibelle(StrToInt(Ident),Libelle);
            TLC.AddChampSupValeur('CHAMP',Prefixe+'_'+Suffixe);
            TLC.AddChampSupValeur('IDENT',Ident);
            TLC.AddChampSupValeur('LETYPE',LeType);
            TLC.AddChampSupValeur('TABLETTE',Tablette);
            TLC.AddChampSupValeur('NUMERO',NumChamp);
            TLC.AddChampSupValeur('LIBCHAMP',Libelle);
            TLC.AddChampSupValeur('COLONNE',Colonne);
            NumChamp := NumChamp + 1;
            LeChamp := ReadTokenpipe(ChampsSelect,';');
       end;
  end
  else
  begin
       Q := OpenSQL('SELECT PAI_PREFIX,PAI_SUFFIX,PAI_IDENT,PAI_LETYPE,PAI_LIBELLE,PAI_LIENASSOC,PAI_COLONNE FROM PAIEPARIM WHERE PAI_LIBELLE<>"" AND PAI_PREFIX="PSA" AND PAI_IDENT<>36 AND PAI_IDENT<>220',True);
       PGTobLesChamps.LoadDetailDB('LesChamps','','',Q,False);
       Ferme(Q);
       for i := 0 to  PGTobLesChamps.Detail.Count -1  do
       begin
            TLC := PGTobLesChamps.Detail[i];
            Prefixe := TLC.GetValue('PAI_PREFIX');
            Suffixe := TLC.GetValue('PAI_SUFFIX');
            Ident := TLC.GetValue('PAI_IDENT');
            LeType := TLC.GetValue('PAI_LETYPE');
            Tablette := TLC.GetValue('PAI_LIENASSOC');
            Libelle := TLC.GetValue('PAI_LIBELLE');
            Colonne := TLC.GetValue('PAI_LIBELLE');
            RendLibelle(StrToInt(Ident),Libelle);
            TLC.AddChampSupValeur('CHAMP',Prefixe+'_'+Suffixe);
            TLC.AddChampSupValeur('IDENT',Ident);
            TLC.AddChampSupValeur('LETYPE',LeType);
            TLC.AddChampSupValeur('TABLETTE',Tablette);
            TLC.AddChampSupValeur('NUMERO',NumChamp);
            TLC.AddChampSupValeur('LIBCHAMP',Libelle);
            TLC.AddChampSupValeur('COLONNE',Colonne);
            NumChamp := NumChamp + 1;
            LeChamp := ReadTokenpipe(ChampsSelect,';');
       end;
  end;
  if Liste = nil then Exit;
  Q_Mul := THQuery(Ecran.FindComponent('Q'));
  if Q_Mul = nil then Exit;
  if ((Liste.nbSelected) > 0) and (not Liste.AllSelected) then
  begin
    InitMoveProgressForm(nil, 'Début du traitement', 'Veuillez patienter SVP ...', Liste.nbSelected, FALSE, TRUE);
    InitMove(Liste.nbSelected, '');
    for i := 0 to Liste.NbSelected - 1 do
    begin
      Liste.GotoLeBOOKMARK(i);
      {$IFDEF EAGLCLIENT}
      TFmul(Ecran).Q.TQ.Seek(Liste.Row - 1);
      {$ENDIF}
      Salarie := TFmul(Ecran).Q.FindField('PSA_SALARIE').asstring;
           TH := Tob.Create('LigneHisto',PGTobHistoSal,-1);
           TH.AddChampSupValeur('SALARIE',Salarie);
           TH.AddChampSupValeur('LIBELLE',rechDom('PGSALARIE',Salarie,False));
           TH.AddChampSupValeur('DATEAPPLIC',DateApplic); //PT4
           TH.AddChampSupValeur('DATEFINVAL',IDate1900);
           TH.AddChampSupValeur('ETAT','#ICO#84');
           For c := 0 to PGTobLesChamps.Detail.count - 1 do
           begin
                IdH := PGTobLesChamps.Detail[c].GetValue('IDENT');
                lEcHAMP := PGTobLesChamps.Detail[c].GetValue('CHAMP');
                LeType := PGTobLesChamps.Detail[c].GetValue('LETYPE');
                Q := OpenSQL('SELECT '+LeChamp+' FROM SALARIES '+
                'WHERE PSA_SALARIE="'+Salarie+'"',True);
                If Not Q.Eof then
                begin
                     If (LeType = 'B') or (LeType='S') or (LeType='T') then
                     begin
                          TH.AddChampSupValeur(LeChamp,Q.FindField(LeChamp).AsString);
                          TH.AddChampSupValeur(LeChamp+'M',Q.FindField(LeChamp).AsString);
                     end
                     else If LeType = 'D' then
                     begin
                          TH.AddChampSupValeur(LeChamp,Q.FindField(LeChamp).AsDateTime);
                          TH.AddChampSupValeur(LeChamp+'M',Q.FindField(LeChamp).AsDateTime);
                     end
                     else if LeType = 'I' then
                     begin
                          TH.AddChampSupValeur(LeChamp,Q.FindField(LeChamp).AsInteger);
                          TH.AddChampSupValeur(LeChamp+'M',Q.FindField(LeChamp).AsInteger);
                     end
                     else If LeType = 'F' then
                     begin
                          TH.AddChampSupValeur(LeChamp,Q.FindField(LeChamp).AsFloat);
                          TH.AddChampSupValeur(LeChamp+'M',Q.FindField(LeChamp).AsFloat);
                     end;
                end
                else
                begin
                     If (LeType = 'B') or (LeType='S') or (LeType='T') then
                     begin
                          TH.AddChampSupValeur(LeChamp,'');
                          TH.AddChampSupValeur(LeChamp+'M','');
                     end
                     else If LeType = 'D' then
                     begin
                          TH.AddChampSupValeur(LeChamp,IDate1900);
                          TH.AddChampSupValeur(LeChamp+'M',IDate1900);
                     end
                     else if (LeType = 'I') or (LeType = 'F') then
                     begin
                          TH.AddChampSupValeur(LeChamp,0);
                          TH.AddChampSupValeur(LeChamp+'M',0);
                     end;
                end;
                Ferme(Q);
                TH.AddChampSupValeur('ORDRE'+IntToStr(c+1),0);
           end;
      TobHisto.Free;
      MoveCurProgressForm(Salarie);
    end;
    FiniMoveProgressForm;
  end
  else if liste.AllSelected then
  begin
    {$IFDEF EAGLCLIENT}
    if (TFMul(Ecran).bSelectAll.Down) then TFMul(Ecran).Fetchlestous; //PT3
    {$ENDIF}
    InitMoveProgressForm(nil, 'Début du traitement', 'Veuillez patienter SVP ...', TFmul(Ecran).Q.RecordCount, FALSE, TRUE);
    InitMove(TFmul(Ecran).Q.RecordCount, '');
    Q_Mul.First;
    while not Q_Mul.EOF do
    begin
      Salarie := TFmul(Ecran).Q.FindField('PSA_SALARIE').asstring;
      TH := Tob.Create('LigneHisto',PGTobHistoSal,-1);
      TH.AddChampSupValeur('SALARIE',Salarie);
      TH.AddChampSupValeur('LIBELLE',rechDom('PGSALARIE',Salarie,False));
      TH.AddChampSupValeur('DATEAPPLIC',DateApplic); //PT4
      TH.AddChampSupValeur('DATEFINVAL',IDate1900);
      TH.AddChampSupValeur('ETAT','#ICO#84');
      For c := 0 to PGTobLesChamps.Detail.count - 1 do
      begin
           IdH := PGTobLesChamps.Detail[c].GetValue('IDENT');
           lEcHAMP := PGTobLesChamps.Detail[c].GetValue('CHAMP');
           LeType := PGTobLesChamps.Detail[c].GetValue('LETYPE');
           Q := OpenSQL('SELECT '+LeChamp+' FROM SALARIES '+
           'WHERE PSA_SALARIE="'+Salarie+'"',True);
           If Not Q.Eof then
           begin
                If (LeType = 'B') or (LeType='S') or (LeType='T') then
                begin
                     TH.AddChampSupValeur(LeChamp,Q.FindField(LeChamp).AsString);
                     TH.AddChampSupValeur(LeChamp+'M',Q.FindField(LeChamp).AsString);
                end
                else If LeType = 'D' then
                begin
                     TH.AddChampSupValeur(LeChamp,Q.FindField(LeChamp).AsDateTime);
                     TH.AddChampSupValeur(LeChamp+'M',Q.FindField(LeChamp).AsDateTime);
                end
                else if LeType = 'I' then
                begin
                     TH.AddChampSupValeur(LeChamp,Q.FindField(LeChamp).AsInteger);
                     TH.AddChampSupValeur(LeChamp+'M',Q.FindField(LeChamp).AsInteger);
                end
                else If LeType = 'F' then
                begin
                     TH.AddChampSupValeur(LeChamp,Q.FindField(LeChamp).AsFloat);
                     TH.AddChampSupValeur(LeChamp+'M',Q.FindField(LeChamp).AsFloat);
                end;
           end
           else
           begin
                If (LeType = 'B') or (LeType='S') or (LeType='T') then
                begin
                     TH.AddChampSupValeur(LeChamp,'');
                     TH.AddChampSupValeur(LeChamp+'M','');
                end
                else If LeType = 'D' then
                begin
                     TH.AddChampSupValeur(LeChamp,IDate1900);
                     TH.AddChampSupValeur(LeChamp+'M',IDate1900);
                end
                else if (LeType = 'I') or (LeType = 'F') then
                begin
                     TH.AddChampSupValeur(LeChamp,0);
                     TH.AddChampSupValeur(LeChamp+'M',0);
                end;
           end;
           Ferme(Q);
           TH.AddChampSupValeur('ORDRE'+IntToStr(c+1),0);
           end;
      Q_Mul.Next;
    end;
    FiniMoveProgressForm;
  end;
end;

procedure TOF_PGMULHISTOGROUPEE.RemplirComboChoix;
VAR Q : TQuery;
    TobListeChamp,TCombo,TC : Tob;
    Num,i : Integer;
    Libelle : String;
    Combo : THMultiValComboBox;
begin
     Q := OpenSQL('SELECT PAI_SUFFIX,PAI_IDENT,PAI_LIBELLE FROM PAIEPARIM WHERE PAI_PREFIX="PSA"'+
     ' AND PAI_LIBELLE <> ""',True);
     TobListeChamp := Tob.Create('LesChamps',Nil,-1);
     TobListeChamp.LoadDetailDB('LesChamps','','',Q,False);
     Ferme(Q);
     TCombo := Tob.Create('RemplirCombo',Nil,-1);
     For i := 0 to TobListeChamp.Detail.Count - 1 do
     begin
          Num := TobListeChamp.Detail[i].GetValue('PAI_IDENT');
          Libelle := TobListeChamp.Detail[i].GetValue('PAI_LIBELLE');
          RendLibelle(Num,Libelle);
          If Libelle <> '' then
          begin
               TC := Tob.Create('FilleRemplirCombo',TCombo,-1);
               TC.AddChampSupValeur('IDENT',Num,False);
               TC.AddChampSupValeur('LIBELLE',Libelle,False);
          end;
     end;
     TobListeChamp.Free;
     Combo := THMultiValComboBox(GetControl('LISTECHAMP'));
     Combo.DataType := '';
     Combo.Items.Clear;
     Combo.Values.Clear;
     For i := 0 to TCombo.Detail.Count - 1 do
     begin
          Combo.Items.Add(TCombo.Detail[i].GetValue('LIBELLE'));
          Combo.Values.Add(TCombo.Detail[i].GetValue('IDENT'));
     end;
     TCombo.Free;
end;

Procedure TOF_PGMULHISTOGROUPEE.RendLibelle(Num : Integer;Var Libelle : String);
begin
     If (Num >=110) and (Num <= 114) then //Salaire mensuel
     begin
          If (Num = 110) and (VH_Paie.PgNbSalLib > 0) then Libelle := 'Salaire mensuel : '+VH_Paie.PgSalLib1
          else if (Num = 111) and (VH_Paie.PgNbSalLib > 1) then Libelle := 'Salaire mensuel : '+VH_Paie.PgSalLib2
          else If (Num = 112) and (VH_Paie.PgNbSalLib > 2) then Libelle := 'Salaire mensuel : '+ VH_Paie.PgSalLib3
          else If (Num = 113) and (VH_Paie.PgNbSalLib > 3) then Libelle := 'Salaire mensuel : '+VH_Paie.PgSalLib4
          else If (Num = 114) and (VH_Paie.PgNbSalLib > 4) then Libelle := 'Salaire mensuel : '+VH_Paie.PgSalLib5
          else Libelle := '';
     end
     else If (Num >=115) and (Num <= 119) then //Salaire ANNUEL
     begin
          If (Num = 115) and (VH_Paie.PgNbSalLib > 0) then Libelle := 'Salaire annuel : '+VH_Paie.PgSalLib1
          else if (Num = 116) and (VH_Paie.PgNbSalLib > 1) then Libelle := 'Salaire annuel : '+VH_Paie.PgSalLib2
          else If (Num = 117) and (VH_Paie.PgNbSalLib > 2) then Libelle := 'Salaire annuel : '+VH_Paie.PgSalLib3
          else If (Num = 118) and (VH_Paie.PgNbSalLib > 3) then Libelle := 'Salaire annuel : '+VH_Paie.PgSalLib4
          else If (Num = 119) and (VH_Paie.PgNbSalLib > 4) then Libelle := 'Salaire annuel : '+VH_Paie.PgSalLib5
          else Libelle := '';
     end
     else If (Num >=157) and (Num <= 160) then // Champs organisation
     begin
          If (Num = 157) and (VH_Paie.PGNbreStatOrg > 0) then Libelle := VH_Paie.PGLibelleOrgStat1
          else if (Num = 158) and (VH_Paie.PGNbreStatOrg > 1) then Libelle := VH_Paie.PGLibelleOrgStat2
          else If (Num = 159) and (VH_Paie.PGNbreStatOrg > 2) then Libelle := VH_Paie.PGLibelleOrgStat3
          else If (Num = 160) and (VH_Paie.PGNbreStatOrg > 3) then Libelle := VH_Paie.PGLibelleOrgStat4
          else Libelle := '';
     end
     else If Num = 161 then //Code statistique
     begin
          If VH_Paie.PGLibCodeStat <> '' then Libelle := VH_Paie.PGLibCodeStat
          else Libelle := '';
     end
     else If (Num >=345) and (Num <= 348) then // Combos libre
     begin
          If (Num = 345) and (VH_Paie.PgNbCombo > 0) then Libelle := VH_Paie.PgLibCombo1
          else if (Num = 346) and (VH_Paie.PgNbCombo > 1) then Libelle := VH_Paie.PgLibCombo2
          else If (Num = 347) and (VH_Paie.PgNbCombo > 2) then Libelle := VH_Paie.PgLibCombo3
          else If (Num = 348) and (VH_Paie.PgNbCombo > 3) then Libelle := VH_Paie.PgLibCombo4
          else Libelle := '';
     end
     else If Num = 49 then Libelle := 'Libellé emploi';
end;

Procedure TOF_PGMULHISTOGROUPEE.NewHisto(Sender : TObject);
var MultiC : THMultivALCombobox;
begin
     MultiC := THMultiValCombobox(GetControl('LISTECHAMP'));
     If MultiC.TEXT = '' then
     begin
          PGIBox('Vous devez choisir au moins une information à modifier',Ecran.Caption);
          Exit;
     end;
     PGTobHistoSal := tob.Create('SaisieHisto',Nil,-1);
     ConstruireTobHistorique;
     AGLLanceFiche('PAY','HISTOGROUPEE','','',GetControlText('LISTECHAMP')+';'+'CREATION');
     Liste.ClearSelected;
     PGTobHistoSal.Free;
     PGTobLesChamps.Free;
end;

procedure TOF_PGMULHISTOGROUPEE.ConstruireTobHistoriqueModif;
var Q_Mul: THQuery;
    Salarie : String;
    Q : TQuery;
    LeChamp,ChampsSelect,IDent,Colonne : String;
    DateDebut,DateFin : TDateTime;
    TH,TLC,TobHisto : Tob;
    c,h,i : integer;
    DHisto : TDateTime;
    IdH,IdS,LeType,Prefixe,Suffixe ,Tablette,Libelle : String;
    NumChamp : Integer;
    MultiC : THMultiValCombobox;
    ExisteHisto : Boolean;
    WhereHisto : String;
begin
  MultiC := THMultiValCombobox(GetControl('LISTECHAMP'));
  NumChamp := 1;
  DateDebut := StrToDate(GetControlText('DATEDEBUT'));
  DateFin := StrToDate(GetControlText('DATEFIN'));
  PGDateDebHisto := DateDebut;
  PGDateFinHisto := DateFin;
  ChampsSelect := GetControlText('LISTECHAMP');
  PGTobLesChamps := Tob.Create('LesChamps',Nil,-1);
  LeChamp := ReadTokenpipe(ChampsSelect,';');
  If not MultiC.Tous then
  begin
       PGTobLesChamps.LoadDetailDB('LesChamps','','',Nil,False);
       While LeChamp <> '' do
       begin
            If Wherehisto <> '' then WhereHisto := WhereHisto + ' OR PHD_PGINFOSMODIF="'+LeChamp+'"'
            else Wherehisto := ' AND (PHD_PGINFOSMODIF="'+LeChamp+'"';
            TLC := Tob.Create('UnChamp',PGTobLesChamps,-1);
            Q := OpenSQL('SELECT PAI_PREFIX,PAI_SUFFIX,PAI_IDENT,PAI_LETYPE,PAI_LIBELLE,PAI_LIENASSOC,PAI_COLONNE FROM PAIEPARIM WHERE PAI_IDENT="'+LeChamp+'"',True);
            If Not Q.Eof then
            begin
                 Prefixe := Q.FindField('PAI_PREFIX').AsString;
                 Suffixe := Q.FindField('PAI_SUFFIX').AsString;
                 Ident := Q.FindField('PAI_IDENT').AsString;
                 LeType := Q.FindField('PAI_LETYPE').AsString;
                 Tablette := Q.FindField('PAI_LIENASSOC').AsString;
                 Libelle := Q.FindField('PAI_LIBELLE').AsString;
                 Colonne := Q.FindField('PAI_COLONNE').AsString;
            end;
            Ferme(Q);
            RendLibelle(StrToInt(Ident),Libelle);
            TLC.AddChampSupValeur('CHAMP',Prefixe+'_'+Suffixe);
            TLC.AddChampSupValeur('IDENT',Ident);
            TLC.AddChampSupValeur('LETYPE',LeType);
            TLC.AddChampSupValeur('TABLETTE',Tablette);
            TLC.AddChampSupValeur('NUMERO',NumChamp);
            TLC.AddChampSupValeur('LIBCHAMP',Libelle);
            TLC.AddChampSupValeur('COLONNE',Colonne);
            NumChamp := NumChamp + 1;
            LeChamp := ReadTokenpipe(ChampsSelect,';');
       end;
  end
  else
  begin
       WhereHisto := '';
       Q := OpenSQL('SELECT PAI_PREFIX,PAI_SUFFIX,PAI_IDENT,PAI_LETYPE,PAI_LIBELLE,PAI_LIENASSOC,PAI_COLONNE FROM PAIEPARIM WHERE PAI_LIBELLE<>"" AND PAI_PREFIX="PSA" AND PAI_IDENT<>36 AND PAI_IDENT<>220',True);
       PGTobLesChamps.LoadDetailDB('LesChamps','','',Q,False);
       Ferme(Q);
       for i := 0 to  PGTobLesChamps.Detail.Count -1  do
       begin
            TLC := PGTobLesChamps.Detail[i];
            Prefixe := TLC.GetValue('PAI_PREFIX');
            Suffixe := TLC.GetValue('PAI_SUFFIX');
            Ident := TLC.GetValue('PAI_IDENT');
            LeType := TLC.GetValue('PAI_LETYPE');
            Tablette := TLC.GetValue('PAI_LIENASSOC');
            Libelle := TLC.GetValue('PAI_LIBELLE');
            Colonne := TLC.GetValue('PAI_LIBELLE');
            RendLibelle(StrToInt(Ident),Libelle);
            TLC.AddChampSupValeur('CHAMP',Prefixe+'_'+Suffixe);
            TLC.AddChampSupValeur('IDENT',Ident);
            TLC.AddChampSupValeur('LETYPE',LeType);
            TLC.AddChampSupValeur('TABLETTE',Tablette);
            TLC.AddChampSupValeur('NUMERO',NumChamp);
            TLC.AddChampSupValeur('LIBCHAMP',Libelle);
            TLC.AddChampSupValeur('COLONNE',Colonne);
            NumChamp := NumChamp + 1;
            LeChamp := ReadTokenpipe(ChampsSelect,';');
       end;
  end;
  If WhereHisto > '' then WhereHisto := WhereHisto+ ')';
  if Liste = nil then Exit;
  Q_Mul := THQuery(Ecran.FindComponent('Q'));
  if Q_Mul = nil then Exit;
  if (Liste.NbSelected = 0) and (TFMul(Ecran).BSelectAll.Down = False) then
  begin
    PGIBox('Aucun élément sélectionné', Ecran.Caption);
    Exit;
  end;
  if ((Liste.nbSelected) > 0) and (not Liste.AllSelected) then
  begin
    InitMoveProgressForm(nil, 'Début du traitement', 'Veuillez patienter SVP ...', Liste.nbSelected, FALSE, TRUE);
    InitMove(Liste.nbSelected, '');
    for i := 0 to Liste.NbSelected - 1 do
    begin
      Liste.GotoLeBOOKMARK(i);
      {$IFDEF EAGLCLIENT}
      TFmul(Ecran).Q.TQ.Seek(Liste.Row - 1);
      {$ENDIF}
      Salarie := TFmul(Ecran).Q.FindField('PSA_SALARIE').asstring;
      TH := Tob.Create('LigneHisto',PGTobHistoSal,-1);
      TH.AddChampSupValeur('SALARIE',Salarie );
      TH.AddChampSupValeur('LIBELLE',Salarie+' '+rechDom('PGSALARIE',Salarie,False));
      TH.AddChampSupValeur('DATEAPPLIC',IDate1900);
      TH.AddChampSupValeur('DATEFINVAL','');
      TH.AddChampSupValeur('ETAT','');
      TH.AddChampSupValeur('TYPELIGNE','S');
      For c := 0 to PGTobLesChamps.Detail.count - 1 do
      begin
           TH.AddChampSupValeur(LeChamp,'');
           TH.AddChampSupValeur(LeChamp+'M','');
           TH.AddChampSupValeur('ORDRE'+IntToStr(c+1),'');
      end;
      Q := OpenSQL('SELECT * FROM PGHISTODETAIL WHERE '+
      'PHD_SALARIE="'+Salarie+'" AND PHD_DATEAPPLIC>="'+UsDateTime(DateDebut)+'" AND '+
      'PHD_DATEAPPLIC<="'+UsDateTime(DateFin)+'" '+WhereHisto,True);
      TobHisto := Tob.Create('Table',Nil,-1);
      TobHisto.LoadDetailDB('Table','','',Q,False);
      Ferme(Q);
      If TobHisto.Detail.Count > 0 then
      begin
           For h := 0 to TobHisto.Detail.Count - 1 do
           begin
                DHisto := TobHisto.Detail[h].GetValue('PHD_DATEAPPLIC');
                TH := PGTobHistoSal.FindFirst(['SALARIE','DATEAPPLIC'],[Salarie,DHisto],False);
                IdS := TobHisto.Detail[h].GetValue('PHD_PGINFOSMODIF');
                If TH = Nil then
                begin
                     TH := Tob.Create('LigneHisto',PGTobHistoSal,-1);
                     TH.AddChampSupValeur('SALARIE',Salarie);
                     TH.AddChampSupValeur('LIBELLE','');
                     TH.AddChampSupValeur('DATEAPPLIC',DHisto);
                     TH.AddChampSupValeur('DATEFINVAL',TobHisto.Detail[h].GetValue('PHD_DATEFINVALID'));
                     If (DHisto < Date) then TH.AddChampSupValeur('ETAT','#ICO#43')
                     else TH.AddChampSupValeur('ETAT','#ICO#66');
                     TH.AddChampSupValeur('TYPELIGNE','D');

                     For c := 0 to PGTobLesChamps.Detail.count - 1 do
                     begin
                          IdH := PGTobLesChamps.Detail[c].GetValue('IDENT');
                          lEcHAMP := PGTobLesChamps.Detail[c].GetValue('CHAMP');
                          LeType := PGTobLesChamps.Detail[c].GetValue('LETYPE');
                          If IdH = IdS then
                          begin
                               If (LeType = 'B') or (LeType='S') or (LeType='T') then
                               begin
                                    TH.AddChampSupValeur(LeChamp,TobHisto.Detail[h].GetValue('PHD_ANCVALEUR'));
                                    TH.AddChampSupValeur(LeChamp+'M',TobHisto.Detail[h].GetValue('PHD_NEWVALEUR'));
                               end
                               else If LeType = 'D' then
                               begin
                                    TH.AddChampSupValeur(LeChamp,StrToDate(TobHisto.Detail[h].GetValue('PHD_ANCVALEUR')));
                                    TH.AddChampSupValeur(LeChamp+'M',StrToDate(TobHisto.Detail[h].GetValue('PHD_NEWVALEUR')));
                               end
                               else if LeType = 'I' then
                               begin
                                    TH.AddChampSupValeur(LeChamp,StrToInt(TobHisto.Detail[h].GetValue('PHD_ANCVALEUR')));
                                    TH.AddChampSupValeur(LeChamp+'M',StrToInt(TobHisto.Detail[h].GetValue('PHD_NEWVALEUR')));
                               end
                               else If LeType = 'F' then
                               begin
                                    TH.AddChampSupValeur(LeChamp,StrToFloat(TobHisto.Detail[h].GetValue('PHD_ANCVALEUR')));
                                    TH.AddChampSupValeur(LeChamp+'M',StrToFloat(TobHisto.Detail[h].GetValue('PHD_NEWVALEUR')));
                               end;
                               TH.AddChampSupValeur('ORDRE'+IntToStr(c+1),TobHisto.Detail[h].GetValue('PHD_ORDRE'));
                          end;
                        {  else
                          begin
                               Q := OpenSQL('SELECT '+LeChamp+' FROM SALARIES '+
                               'WHERE PSA_SALARIE="'+Salarie+'"',True);
                               If Not Q.Eof then
                               begin
                                    If (LeType = 'B') or (LeType='S') or (LeType='T') then
                                    begin
                                         TH.AddChampSupValeur(LeChamp,Q.FindField(LeChamp).AsString);
                                         TH.AddChampSupValeur(LeChamp+'M',Q.FindField(LeChamp).AsString);
                                    end
                                    else If LeType = 'D' then
                                    begin
                                         TH.AddChampSupValeur(LeChamp,Q.FindField(LeChamp).AsDateTime);
                                         TH.AddChampSupValeur(LeChamp+'M',Q.FindField(LeChamp).AsDateTime);
                                    end
                                    else if LeType = 'I' then
                                    begin
                                         TH.AddChampSupValeur(LeChamp,Q.FindField(LeChamp).AsInteger);
                                         TH.AddChampSupValeur(LeChamp+'M',Q.FindField(LeChamp).AsInteger);
                                    end
                                    else If LeType = 'F' then
                                    begin
                                         TH.AddChampSupValeur(LeChamp,Q.FindField(LeChamp).AsFloat);
                                         TH.AddChampSupValeur(LeChamp+'M',Q.FindField(LeChamp).AsFloat);
                                    end;
                               end
                               else
                               begin
                                    If (LeType = 'B') or (LeType='S') or (LeType='T') then
                                    begin
                                         TH.AddChampSupValeur(LeChamp,'');
                                         TH.AddChampSupValeur(LeChamp+'M','');
                                    end
                                    else If LeType = 'D' then
                                    begin
                                         TH.AddChampSupValeur(LeChamp,IDate1900);
                                         TH.AddChampSupValeur(LeChamp+'M',IDate1900);
                                    end
                                    else if (LeType = 'I') or (LeType = 'F') then
                                    begin
                                         TH.AddChampSupValeur(LeChamp,0);
                                         TH.AddChampSupValeur(LeChamp+'M',0);
                                    end;
                               end;
                               Ferme(Q);
                               TH.AddChampSupValeur('ORDRE'+IntToStr(c+1),0);
                          end;}
                     end;
                end
                else
                begin
                     For c := 0 to PGTobLesChamps.Detail.count - 1 do
                     begin
                          IdH := PGTobLesChamps.Detail[c].GetValue('IDENT');
                          lEcHAMP := PGTobLesChamps.Detail[c].GetValue('CHAMP');
                          LeType := PGTobLesChamps.Detail[c].GetValue('LETYPE');
                          If IdH = IdS then
                          begin
                               If (LeType = 'B') or (LeType='S') or (LeType='T') then
                               begin
                                    TH.PutValue(LeChamp,TobHisto.Detail[h].GetValue('PHD_ANCVALEUR'));
                                    TH.PutValue(LeChamp+'M',TobHisto.Detail[h].GetValue('PHD_NEWVALEUR'));
                               end
                               else If LeType = 'D' then
                               begin
                                    TH.PutValue(LeChamp,StrToDate(TobHisto.Detail[h].GetValue('PHD_ANCVALEUR')));
                                    TH.PutValue(LeChamp+'M',StrToDate(TobHisto.Detail[h].GetValue('PHD_NEWVALEUR')));
                               end
                               else if LeType = 'I' then
                               begin
                                    TH.PutValue(LeChamp,StrToInt(TobHisto.Detail[h].GetValue('PHD_ANCVALEUR')));
                                    TH.PutValue(LeChamp+'M',StrToInt(TobHisto.Detail[h].GetValue('PHD_NEWVALEUR')));
                               end
                               else If LeType = 'F' then
                               begin
                                    TH.PutValue(LeChamp,StrToFloat(TobHisto.Detail[h].GetValue('PHD_ANCVALEUR')));
                                    TH.PutValue(LeChamp+'M',StrToFloat(TobHisto.Detail[h].GetValue('PHD_NEWVALEUR')));
                               end;
                               TH.AddChampSupValeur('ORDRE'+IntToStr(c+1),TobHisto.Detail[h].GetValue('PHD_ORDRE'));
                          end;
                     end;
                end;
           end;
      end;
      TobHisto.Free;
      MoveCurProgressForm(Salarie);
    end;
    FiniMoveProgressForm;
  end
  else if liste.AllSelected then
  begin
         {$IFDEF EAGLCLIENT}
     if (TFMul(Ecran).bSelectAll.Down) then TFMul(Ecran).Fetchlestous;
    {$ENDIF}
    InitMoveProgressForm(nil, 'Début du traitement', 'Veuillez patienter SVP ...', TFmul(Ecran).Q.RecordCount, FALSE, TRUE);
    InitMove(TFmul(Ecran).Q.RecordCount, '');
    Q_Mul.First;
    while not Q_Mul.EOF do
    begin
      Salarie := TFmul(Ecran).Q.FindField('PSA_SALARIE').asstring;
      TH := Tob.Create('LigneHisto',PGTobHistoSal,-1);
      TH.AddChampSupValeur('SALARIE',Salarie);
      TH.AddChampSupValeur('LIBELLE',rechDom('PGSALARIE',Salarie,False));
      TH.AddChampSupValeur('DATEAPPLIC',DateFin);
      TH.AddChampSupValeur('DATEFINVAL',IDate1900);
      TH.AddChampSupValeur('ETAT','');
      For c := 0 to PGTobLesChamps.Detail.count - 1 do
      begin
           IdH := PGTobLesChamps.Detail[c].GetValue('IDENT');
           lEcHAMP := PGTobLesChamps.Detail[c].GetValue('CHAMP');
           LeType := PGTobLesChamps.Detail[c].GetValue('LETYPE');
           Q := OpenSQL('SELECT '+LeChamp+' FROM SALARIES '+
           'WHERE PSA_SALARIE="'+Salarie+'"',True);
           If Not Q.Eof then
           begin
                If (LeType = 'B') or (LeType='S') or (LeType='T') then
                begin
                     TH.AddChampSupValeur(LeChamp,Q.FindField(LeChamp).AsString);
                     TH.AddChampSupValeur(LeChamp+'M',Q.FindField(LeChamp).AsString);
                end
                else If LeType = 'D' then
                begin
                     TH.AddChampSupValeur(LeChamp,Q.FindField(LeChamp).AsDateTime);
                     TH.AddChampSupValeur(LeChamp+'M',Q.FindField(LeChamp).AsDateTime);
                end
                else if LeType = 'I' then
                begin
                     TH.AddChampSupValeur(LeChamp,Q.FindField(LeChamp).AsInteger);
                     TH.AddChampSupValeur(LeChamp+'M',Q.FindField(LeChamp).AsInteger);
                end
                else If LeType = 'F' then
                begin
                     TH.AddChampSupValeur(LeChamp,Q.FindField(LeChamp).AsFloat);
                     TH.AddChampSupValeur(LeChamp+'M',Q.FindField(LeChamp).AsFloat);
                end;
           end
           else
           begin
                If (LeType = 'B') or (LeType='S') or (LeType='T') then
                begin
                     TH.AddChampSupValeur(LeChamp,'');
                     TH.AddChampSupValeur(LeChamp+'M','');
                end
                else If LeType = 'D' then
                begin
                     TH.AddChampSupValeur(LeChamp,IDate1900);
                     TH.AddChampSupValeur(LeChamp+'M',IDate1900);
                end
                else if (LeType = 'I') or (LeType = 'F') then
                begin
                     TH.AddChampSupValeur(LeChamp,0);
                     TH.AddChampSupValeur(LeChamp+'M',0);
                end;
           end;
           Ferme(Q);
           TH.AddChampSupValeur('ORDRE'+IntToStr(c+1),0);
           end;
      Q_Mul.Next;
    end;
    FiniMoveProgressForm;
  end;
end;

//PT5
procedure TOF_PGMULHISTOGROUPEE.OnActivate(Sender: TObject);
begin
  inherited;
  if AutoValide then
  begin
    TFMul(Ecran).BSelectAll.Click;
    Liste.AllSelected := True;
    (GetControl('BOuvrir')    as TToolbarButton97).OnClick(Sender);
  end;
end;

Initialization
registerclasses([TOF_PGMULHISTOGROUPEE]) ;

end.




