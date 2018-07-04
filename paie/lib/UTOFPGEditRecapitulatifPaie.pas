{***********UNITE*************************************************
Auteur  ...... : Paie PGI
Créé le ...... : 09/05/2003
Modifié le ... : 09/05/2003
Description .. : Source TOF de la FICHE : RECAPPAIE_ETAT ()
Mots clefs ... : PAIE;EDITION
*****************************************************************}
{PT1 12/09/2003 SB V42 FQ 10799 Intégration des ruptures
 PT2 02/10/2003 SB V_42 Affichage des ongles si gestion paramsoc des combos libres
 PT3 05/05/2004 SB V_50 FQ 11222 Ajout des taux salariales et patronales
 PT4 13/05/2004 SB V_50 FQ 11150 Editin des rubriques non imprimables ou non sur coche
 PT5 14/05/2004 SB V_50 FQ 11197 Traitements Etats chainés
 PT6 01/12/2005 SB V_65 FQ 12734 Maj bulletin ordreetat
 PT7 15/05/2006 RM V_7  FQ 11486 Renseigner les bases des rémunérations en indiquant le type (H,Q,J,M)
 PT8 28/06/2007 FC V_72 FQ 13508 Ajout case salariés sortis
 PT9 17/09/2007 FC V_80 FQ 14641 Rubriques 4500 et 4500.R sur une seule ligne
 }
Unit UTOFPGEditRecapitulatifPaie ;

Interface

Uses StdCtrls,Classes,
{$IFNDEF EAGLCLIENT}
     QRS1,
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}  //PT7
{$ELSE}
     EQRS1,
     UtileAgl,  //PT7
{$ENDIF}
     sysutils,ComCtrls,HCtrls,HEnt1,HQry,UTOF,paramsoc,UTOFPGEtats,Ed_tools,
     UTOB ,UtobDebug ;  //PT7
Type
  TOF_PGRECAPPAIE_ETAT = Class (TOF_PGEtats)  { PT5 }
    procedure OnArgument (S : String ) ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override;
    Procedure OnClose                  ; override; //PT7
  private
    AjoutChamp,Origine : String;
    DateDebut, DateFin: TDateTime; { PT5 }
    TSQL : Tob;     //PT7
    Procedure ExitEdit ( Sender : TObject ) ;
    Procedure ClickRupture ( Sender : TObject ) ;
  end ;

Implementation

uses PGOutils2,EntPaie,PGEditOutils,PGEditOutils2,PgEdtEtat;

{ TOF_PGRECAPPAIE_ETAT }

procedure TOF_PGRECAPPAIE_ETAT.ExitEdit(Sender: TObject);
var edit : thedit;
begin
  edit:=THEdit(Sender);
  if edit <> nil then
    if edit.text<>'' then
      if (VH_Paie.PgTypeNumSal='NUM') and (length(Edit.text)<11) and (isnumeric(edit.text)) then
        edit.text:=AffectDefautCode(edit,10);
end;

procedure TOF_PGRECAPPAIE_ETAT.OnArgument(S: String);
var
  Defaut                               : THEdit;
  Min,Max,DebPer,FinPer,ExerPerEncours : string;
  Check                                : TCheckBox;
begin
  inherited;

  { DEB PT5 }
  DateDebut := idate1900;
  DateFin := idate1900;
  if (S <> '') then
  begin
    if (Pos('CHAINES', S) > 0) then
    begin
      Origine := ReadTokenSt(S);
      if trim(S) <> '' then
      begin
        DateDebut := StrToDate(ReadTokenSt(S));
        DateFin := StrToDate(ReadTokenSt(S));
      end;
    end;
  end
  else Origine := 'MENU';
  { FIN PT5 }

VisibiliteChamp (Ecran);
VisibiliteChampLibre (Ecran);
{ DEB PT2 }
SetControlProperty('TBCOMPLEMENT','Tabvisible',(VH_Paie.PGNbreStatOrg>0) OR (VH_Paie.PGLibCodeStat<>''));
SetControlProperty('TBCHAMPLIBRE','Tabvisible',(VH_Paie.PgNbCombo>0));
{ FIN PT2 }
{ Valeur par défaut }
RecupMinMaxTablette('PG','SALARIES','PSA_SALARIE',Min,Max);
Defaut:=ThEdit(getcontrol('PHB_SALARIE'));
If Defaut<>nil then Begin Defaut.text:=Min; Defaut.OnExit:=ExitEdit; End;
Defaut:=ThEdit(getcontrol('PHB_SALARIE_'));
If Defaut<>nil then Begin Defaut.text:=Max; Defaut.OnExit:=ExitEdit;  End;
RecupMinMaxTablette('PG','ETABLISS','ET_ETABLISSEMENT',Min,Max);
Defaut:=ThEdit(getcontrol('PHB_ETABLISSEMENT'));
If Defaut<>nil then Begin Defaut.text:=Min;  End;
Defaut:=ThEdit(getcontrol('PHB_ETABLISSEMENT_'));
If Defaut<>nil then Begin Defaut.text:=Max; End;

if RendPeriodeEnCours (ExerPerEncours,DebPer,FinPer)=True then
   begin
   SetControlText('PHB_DATEDEBUT',DebPer);
   SetControlText('PHB_DATEFIN'  ,FinPer);
   end;

   // Evenement ONCHANGE
{ DEB PT1 }
Check:=TCheckBox(GetControl('CN1'));
if Check<>nil then Check.OnClick := ClickRupture;
Check:=TCheckBox(GetControl('CN2'));
if Check<>nil then Check.OnClick := ClickRupture;
Check:=TCheckBox(GetControl('CN3'));
if Check<>nil then Check.OnClick := ClickRupture;
Check:=TCheckBox(GetControl('CN4'));
if Check<>nil then Check.OnClick := ClickRupture;
Check:=TCheckBox(GetControl('CN5'));
if Check<>nil then Check.OnClick := ClickRupture;
Check:=TCheckBox(GetControl('CETAB'));
if Check<>nil then Check.OnClick := ClickRupture;
Check:=TCheckBox(GetControl('CL1'));
if Check<>nil then Check.OnClick := ClickRupture;
Check:=TCheckBox(GetControl('CL2'));
if Check<>nil then Check.OnClick := ClickRupture;
Check:=TCheckBox(GetControl('CL3'));
if Check<>nil then Check.OnClick := ClickRupture;
Check:=TCheckBox(GetControl('CL4'));
if Check<>nil then Check.OnClick := ClickRupture;
{ FIN PT1 }
SetControlText('DOSSIER',GetParamSoc ('SO_LIBELLE'));

//DEB PT5 Affect critère standard
if (origine = 'CHAINES') and (DateFin <> idate1900) then
begin
SetControlText('PHB_DATEDEBUT',DateToStr(DateDebut));
SetControlText('PHB_DATEFIN'  ,DateToStr(DateFin));
end;
//FIN PT5

end;

{ DEB PT1 }
procedure TOF_PGRECAPPAIE_ETAT.ClickRupture(Sender: TObject);
var
CEtab,CN1,CN2,CN3,CN4,CN5,CL1,CL2,CL3,CL4:TCheckBox;
BEGIN
BloqueChampLibre(Ecran);
RecupChampRupture(Ecran);
AjoutChamp:='';
//Gestion des ruptures
CEtab:=TCheckBox(GetControl('CETAB'));
if (CEtab<>nil) then
  if (CEtab.Checked=True) then  AjoutChamp:='PHB_ETABLISSEMENT';
CN1:=TCheckBox(GetControl('CN1'));
CN2:=TCheckBox(GetControl('CN2'));
CN3:=TCheckBox(GetControl('CN3'));
CN4:=TCheckBox(GetControl('CN4'));
CN5:=TCheckBox(GetControl('CN5'));
if (CN1<>nil)and(CN2<>nil)and(CN3<>nil)and(CN4<>nil)and(CN5<>nil) then
   begin
   if (CN1.Checked=True) then AjoutChamp:='PHB_TRAVAILN1';
   if (CN2.Checked=True) then AjoutChamp:='PHB_TRAVAILN2';
   if (CN3.Checked=True) then AjoutChamp:='PHB_TRAVAILN3';
   if (CN4.Checked=True) then AjoutChamp:='PHB_TRAVAILN4';
   if (CN5.Checked=True) then AjoutChamp:='PHB_CODESTAT' ;
   end;

CL1:=TCheckBox(GetControl('CL1'));
CL2:=TCheckBox(GetControl('CL2'));
CL3:=TCheckBox(GetControl('CL3'));
CL4:=TCheckBox(GetControl('CL4'));
if (CL1<>nil)and(CL2<>nil)and(CL3<>nil)and(CL4<>nil)then
   begin
   if (CL1.Checked=True) then AjoutChamp:='PHB_LIBREPCMB1';
   if (CL2.Checked=True) then AjoutChamp:='PHB_LIBREPCMB2';
   if (CL3.Checked=True) then AjoutChamp:='PHB_LIBREPCMB3';
   if (CL4.Checked=True) then AjoutChamp:='PHB_LIBREPCMB4';
   End;

SetControlText('XX_RUPTURE1',AjoutChamp);
SetControlText('CHAMPRUPT',ConvertPrefixe(AjoutChamp,'PHB','PPU'));
If AjoutChamp<>'' then AjoutChamp:=AjoutChamp+', ';
end;
{ FIN PT1 }


procedure TOF_PGRECAPPAIE_ETAT.OnUpdate;
Var
StWhere,Tempo,StRubimpr : String;
X                       : Integer;
SQL,Crupture  : String;  //PT7
Q             : TQuery;  //PT7
TCPLT         : Tob;     //PT7
i ,y          : Integer; //PT7
Cfait         : Boolean; //PT7

begin
  inherited;
{ DEB PT6 }
if GetParamSocSecur('SO_PGMAJBULLETIN',False) = False then
Begin
InitMoveProgressForm(nil,'Vérification des données','Veuillez patienter SVP',10,False,True);
ExecuteSql('UPDATE HISTOBULLETIN SET PHB_ORDREETAT=3 WHERE PHB_NATURERUB<>"AAA" AND PHB_ORDREETAT<>3 ');
ExecuteSql('UPDATE HISTOANALPAIE SET PHA_ORDREETAT=3 WHERE PHA_NATURERUB<>"AAA" AND PHA_ORDREETAT<>3 ');
SetParamSoc('SO_PGMAJBULLETIN','X');
MoveCurProgressForm('');
ChargeParamsPaie;
FiniMoveProgressForm;
End;
{ FIN PT6 }

{ DEB PT1 }
StWhere:=RecupWhereCritere(TPageControl(GetControl('Pages')));
tempo:='';
x:=Pos('(',StWhere);
if x>0 then Tempo:=copy(StWhere,x,(Length(StWhere)-5));
if tempo<>'' then StWhere:='AND '+Tempo;

  { DEB PT8 }
  if (GetControltext('CKSORTIE')='X') and (UsdateTime(StrToDate(GetControlText('PHB_DATEFIN')))<>UsdateTime(idate1900)) then
    StWhere:=StWhere+' AND PHB_SALARIE IN (SELECT PSA_SALARIE FROM SALARIES WHERE PSA_DATESORTIE IS NULL '+
                      'OR PSA_DATESORTIE="'+Usdatetime(idate1900)+'" '+
                      'OR PSA_DATESORTIE>"'+UsdateTime(StrToDate(GetControlText('PHB_DATEFIN')))+'")';
  { FIN PT8 }

SetControlText('STWHERE',ConvertPrefixe(StWhere,'PHB','PPU'));

if GetControlText('CKRUBNONIMP') = '-' then StRubimpr := 'AND PHB_IMPRIMABLE="X" ' else StRubimpr := ''; { PT4 }
//PT7 TFQRS1(Ecran).WhereSQL:='SELECT PHB_ORDREETAT,PHB_RUBRIQUE,PHB_NATURERUB, '+
SQL:='SELECT PHB_ORDREETAT,PHB_RUBRIQUE,PHB_COTREGUL,PHB_NATURERUB, '+                   //PT9
'PHB_LIBELLE,PHB_SENSBUL,'+AjoutChamp+'COUNT(DISTINCT PHB_SALARIE) AS NBSAL,SUM(PHB_MTREM) AS MTREM , '+
'SUM(PHB_BASECOT) AS BASECOT ,PHB_TAUXSALARIAL,SUM(PHB_MTSALARIAL)AS MTSAL , '+ { PT3 }
'PHB_TAUXPATRONAL,SUM(PHB_MTPATRONAL) AS MTPAT '+ { PT3 }
'FROM HISTOBULLETIN '+
'WHERE PHB_NATURERUB<>"BAS" AND PHB_RUBRIQUE NOT LIKE "%.%" '+
'AND (PHB_MTSALARIAL<>0 OR PHB_MTPATRONAL<>0 OR PHB_MTREM<>0) '+StRubimpr+StWhere+ { PT4 }
'GROUP BY '+AjoutChamp+'PHB_ORDREETAT,PHB_NATURERUB,PHB_RUBRIQUE,PHB_COTREGUL,PHB_TAUXSALARIAL,PHB_TAUXPATRONAL,PHB_LIBELLE,PHB_SENSBUL '+ { PT3 } //PT9
'ORDER BY '+AjoutChamp+'PHB_ORDREETAT,PHB_NATURERUB,PHB_RUBRIQUE,PHB_COTREGUL'; //PT9
{ FIN PT1 }
{ DEB PT7 }
  FreeAndNil (TSQL);
  Q := OpenSQL(SQL,True);
  TSQL := Tob.Create('Donnees',Nil,-1);
  TSQL.LoadDetailDB('Donnees','','',Q,False);
  Ferme(Q);

SQL:='SELECT PHB_RUBRIQUE,PHB_COTREGUL,'+ AjoutChamp + 'SUM(PHB_BASEREM) AS BASEREM,PRM_BASEMTQTE,CO_LIBELLE '+  //PT9
'FROM HISTOBULLETIN '+
'LEFT JOIN REMUNERATION ON ##PRM_PREDEFINI## (PHB_RUBRIQUE = PRM_RUBRIQUE) ' +
'LEFT JOIN COMMUN ON (PRM_BASEMTQTE = CO_CODE) ' +
'WHERE PHB_ORDREETAT = "1" And PHB_NATURERUB = "AAA" AND PHB_RUBRIQUE NOT LIKE "%.%" And PHB_BASEREM <> 0 ' +
'AND CO_TYPE = "PQB" '+StRubimpr+StWhere+
'GROUP BY '+AjoutChamp+'PHB_RUBRIQUE,PHB_COTREGUL,PRM_BASEMTQTE,CO_LIBELLE '+ //PT9
'ORDER BY '+AjoutChamp+'PHB_RUBRIQUE,PHB_COTREGUL';   //PT9

  Q := OpenSQL(SQL,True);
  TCPLT := Tob.Create('Complement',Nil,-1);
  TCPLT.LoadDetailDB('Complement','','',Q,False);
  Ferme(Q);

  Crupture := Trim(GETControlText('XX_RUPTURE1'));

  For i := 0 To TSQL.Detail.Count -1 do
  begin
     If Crupture <> '' Then
        TSQL.Detail[i].AddChampSupValeur('RUPTURE1',TSQL.Detail[i].GetValue(Crupture))
     Else
        TSQL.Detail[i].AddChampSupValeur('RUPTURE1','');

     Cfait := False;
     If (TSQL.Detail[i].GetValue('PHB_ORDREETAT')='1') Then
     Begin
        For y := 0 To TCPLT.Detail.Count -1 do
        Begin
           If (TCPLT.Detail[y].GetValue('PHB_RUBRIQUE') = TSQL.Detail[i].GetValue('PHB_RUBRIQUE')) And
             ((Crupture = '') Or (TCPLT.Detail[y].GetValue(Crupture) = TSQL.Detail[i].GetValue(Crupture))) Then
           Begin
              TSQL.Detail[i].PutValue('BASECOT',TCPLT.Detail[y].GetValue('BASEREM'));
              TSQL.Detail[i].AddChampSupValeur('TypeBase',TCPLT.Detail[y].GetValue('CO_LIBELLE'));
              Cfait := True;
              Break;
           End;
        End;
     End;
     If Cfait = False Then TSQL.Detail[i].AddChampSupValeur('TypeBase','');
     //DEB PT9
     //Quand on a une rubrique de régul, on a pas le .R dans la rubrique donc se servir du champ PHB_COTREGUL pour la rupture
     if TSQL.Detail[i].GetValue('PHB_COTREGUL') = 'REG' then
       TSQL.Detail[i].PutValue('PHB_COTREGUL',TSQL.Detail[i].GetValue('PHB_RUBRIQUE') + '.R')
     else
       TSQL.Detail[i].PutValue('PHB_COTREGUL',TSQL.Detail[i].GetValue('PHB_RUBRIQUE'));
     //FIN PT9
  End;
  //TOBDebug(TCPLT);
  //TOBDebug(TSQL);
  TCPLT.Free;
  TFQRS1(Ecran).LaTob:= TSQL;
{ FIN PT7 }
end;

procedure TOF_PGRECAPPAIE_ETAT.OnLoad;
begin
  inherited;
  //DEB PT5
  if (Origine = 'CHAINES') and (DateDebut > idate1900) then
  begin
  SetControlText('PHB_DATEDEBUT',DateToStr(DateDebut));
  SetControlText('PHB_DATEFIN'  ,DateToStr(DateFin));
  end;
  //FIN PT5
end;

{ DEB PT7 }
Procedure TOF_PGRECAPPAIE_ETAT.OnClose ;
Begin
  Inherited ;
  FreeAndNil (TSQL);
End ;
{ FIN PT7 }

Initialization
  registerclasses ( [ TOF_PGRECAPPAIE_ETAT ] ) ;
end.
