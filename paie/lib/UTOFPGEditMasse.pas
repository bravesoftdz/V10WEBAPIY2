/////////////// SOURCE NON UTILISE

unit UTOFPGEditMasse;

interface
uses StdCtrls,Controls,Classes,Graphics,forms,sysutils,ComCtrls,
{$IFDEF EAGLCLIENT}
     eQRS1, UtileAGL,MaineAgl,
{$ELSE}
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}QRS1,Fe_main,EdtEtat,edtRetat,
{$ENDIF}
     HCtrls,HEnt1,HMsgBox,UTOF,ParamDat,PGIenv,
     ParamSoc,HQry;

Type
     TOF_PGEDITMASSE_ETAT = Class (TOF)
       public
       procedure OnArgument(Arguments : String ) ; override ;
       procedure OnUpdate; override;
       private
       AjoutChamp,ChampOrderBy : string;
       procedure DateElipsisclick(Sender: TObject);
       procedure MonnaieInverse(Sender: TObject);
       procedure ChangeLieuTravail(Sender : TObject);
       procedure ExitEdit(Sender: TObject);
       Procedure OnClickBulletin(Sender: TObject);
       Procedure OnClickChargesSoc(Sender: TObject);
       Procedure OnClickReglement(Sender: TObject);
       Procedure ForceLibellePaie(Sender: TObject);
       function  RecupCritereGeneral(Prefixe : String) : string;
     END ;
implementation

uses PgEditOutils,EntPaie,PgOutils,PGEdtEtat,UTOB;


{ TOF_PGEDITMASSE_ETAT }
procedure TOF_PGEDITMASSE_ETAT.OnArgument(Arguments: String);
var
Check :TCheckBox;
CDd, CDf,Defaut : THEdit;
Min,Max,DebPer,FinPer,ExerPerEncours:string;
Q : TQuery;
begin
  inherited;
Defaut:=ThEdit(getcontrol('DOSSIER'));
If Defaut<>nil then
  Defaut.text:=GetParamSoc ('SO_LIBELLE');
VisibiliteChamp (Ecran);
VisibiliteChampLibre (Ecran);
//Valeur par défaut
RecupMinMaxTablette('PG','SALARIES','PSA_SALARIE',Min,Max);
Defaut:=ThEdit(getcontrol('PPU_SALARIE'));
If Defaut<>nil then Begin Defaut.text:=Min;  Defaut.OnExit:=ExitEdit; end;
Defaut:=ThEdit(getcontrol('PPU_SALARIE_'));
If Defaut<>nil then Begin Defaut.text:=Max;   Defaut.OnExit:=ExitEdit; end;
RecupMinMaxTablette('PG','ETABLISS','ET_ETABLISSEMENT',Min,Max);
Defaut:=ThEdit(getcontrol('PPU_ETABLISSEMENT'));
If Defaut<>nil then Defaut.text:=Min;
Defaut:=ThEdit(getcontrol('PPU_ETABLISSEMENT_'));
If Defaut<>nil then Defaut.text:=Max;
CDd:=THEdit(GetControl ('XX_VARIABLEDEB'));
CDf:=THEdit(GetControl ('XX_VARIABLEFIN'));
If  CDd<>NIL then CDd.OnElipsisClick:=DateElipsisclick;
If  CDf<>NIL then  CDf.OnElipsisClick:=DateElipsisclick;
if RendPeriodeEnCours (ExerPerEncours,DebPer,FinPer)=True then
   begin
   If  CDd<>NIL then CDd.text:=DebPer;
   If  CDf<>NIL then CDf.text:=FinPer;
   end;

Q:=OpenSql('SELECT D_PARITEEURO FROM DEVISE WHERE D_DEVISE="'+VH_Paie.PGMonnaieTenue+'"',True);
if not Q.eof then
  SetControlText('TAUX',Q.FindField('D_PARITEEURO').asstring)
else
  SetControlText('TAUX','1');
Ferme(Q);

SetControlChecked('CKEURO',VH_Paie.PGTenueEuro);


    //Evenement ONCHANGE
Check:=TCheckBox(GetControl('CALPHA'));
if Check<>nil then Check.OnClick:=ChangeLieuTravail;
Check:=TCheckBox(GetControl('CN1'));
if Check<>nil then  Check.OnClick:=ChangeLieuTravail;
Check:=TCheckBox(GetControl('CN2'));
if Check<>nil then  Check.OnClick:=ChangeLieuTravail;
Check:=TCheckBox(GetControl('CN3'));
if Check<>nil then Check.OnClick:=ChangeLieuTravail;
Check:=TCheckBox(GetControl('CN4'));
if Check<>nil then Check.OnClick:=ChangeLieuTravail;
Check:=TCheckBox(GetControl('CN5'));
if Check<>nil then Check.OnClick:=ChangeLieuTravail;
Check:=TCheckBox(GetControl('CETAB'));
if Check<>nil then Check.OnClick:=ChangeLieuTravail;
Check:=TCheckBox(GetControl('CL1'));
if Check<>nil then Check.OnClick:=ChangeLieuTravail;
Check:=TCheckBox(GetControl('CL2'));
if Check<>nil then Check.OnClick:=ChangeLieuTravail;
Check:=TCheckBox(GetControl('CL3'));
if Check<>nil then Check.OnClick:=ChangeLieuTravail;
Check:=TCheckBox(GetControl('CL4'));
if Check<>nil then Check.OnClick:=ChangeLieuTravail;
Check := TCheckBox(GetControl('CHMONNAIEINV'));
if Check<>nil then begin Check.OnClick:=MonnaieInverse; PgMonnaieInv:=Check.checked; End;
Check := TCheckBox(GetControl('CKEURO'));
if Check<>nil then Check.Checked:=VH_Paie.PGTenueEuro;
Check:=TCheckBox(GetControl('CKBULLETIN'));
if Check<>nil then Check.OnClick:=OnClickBulletin;
Check:=TCheckBox(GetControl('CKFORCELIB'));
if Check<>nil then Check.OnClick:=ForceLibellePaie;
Check:=TCheckBox(GetControl('CKCHARGESOC'));
if Check<>nil then Check.OnClick:=OnClickChargesSoc;
Check:=TCheckBox(GetControl('CKREGLEMENT'));
if Check<>nil then Check.OnClick:=OnClickReglement;
end;


procedure TOF_PGEDITMASSE_ETAT.ChangeLieuTravail(Sender: TObject);
var
CEtab,CN1,CN2,CN3,CN4,CN5,Alpha,CL1,CL2,CL3,CL4:TCheckBox;
begin
BloqueChampLibre(Ecran);
RecupChampRupture(Ecran);

CEtab:=TCheckBox(GetControl('CETAB'));
if (CEtab<>nil) then
   if (CEtab.Checked=True) then
     begin AjoutChamp:='PPU_ETABLISSEMENT, ';ChampOrderBy:='PPU_ETABLISSEMENT, ';
     PgChampRupt:='PPU_ETABLISSEMENT';
     end;
CN1:=TCheckBox(GetControl('CN1'));
CN2:=TCheckBox(GetControl('CN2'));
CN3:=TCheckBox(GetControl('CN3'));
CN4:=TCheckBox(GetControl('CN4'));
CN5:=TCheckBox(GetControl('CN5'));
Alpha:=TCheckBox(GetControl('CALPHA'));

if (CN1<>nil)and(CN2<>nil)and(CN3<>nil)and(CN4<>nil)and(CN5<>nil) then
   begin
   if (CN1.Checked=True) then
       begin AjoutChamp:='PPU_TRAVAILN1, ';ChampOrderBy:='PPU_TRAVAILN1, '; PgChampRupt:='PPU_TRAVAILN1';end;
   if (CN2.Checked=True) then
       begin AjoutChamp:='PPU_TRAVAILN2, ';ChampOrderBy:='PPU_TRAVAILN2, ';PgChampRupt:='PPU_TRAVAILN2'; end;
   if (CN3.Checked=True) then
       begin AjoutChamp:='PPU_TRAVAILN3, ';ChampOrderBy:='PPU_TRAVAILN3, '; PgChampRupt:='PPU_TRAVAILN3';end;
   if (CN4.Checked=True) then
       begin AjoutChamp:='PPU_TRAVAILN4, ';ChampOrderBy:='PPU_TRAVAILN4, '; PgChampRupt:='PPU_TRAVAILN4';end;
   if (CN5.Checked=True) then
       begin AjoutChamp:='PPU_CODESTAT, ';ChampOrderBy:='PPU_CODESTAT, ';  PgChampRupt:='PPU_CODESTAT'; end;
   end;

CL1:=TCheckBox(GetControl('CL1'));
CL2:=TCheckBox(GetControl('CL2'));
CL3:=TCheckBox(GetControl('CL3'));
CL4:=TCheckBox(GetControl('CL4'));
if (CL1<>nil)and(CL2<>nil)and(CL3<>nil)and(CL4<>nil)then
   begin
   if (CL1.Checked=True) then
       begin AjoutChamp:='PPU_LIBREPCMB1, ';ChampOrderBy:='PPU_LIBREPCMB1, '; PgChampRupt:='PPU_LIBREPCMB1';end;
   if (CL2.Checked=True) then
       begin AjoutChamp:='PPU_LIBREPCMB2, ';ChampOrderBy:='PPU_LIBREPCMB2, ';PgChampRupt:='PPU_LIBREPCMB2'; end;
   if (CL3.Checked=True) then
       begin AjoutChamp:='PPU_LIBREPCMB3, ';ChampOrderBy:='PPU_LIBREPCMB3, '; PgChampRupt:='PPU_LIBREPCMB3';end;
   if (CL4.Checked=True) then
       begin AjoutChamp:='PPU_LIBREPCMB4, ';ChampOrderBy:='PPU_LIBREPCMB4, '; PgChampRupt:='PPU_LIBREPCMB4';end;
   End;

    // TRI SUR LIBELLE SALARIE + CHAMPS RUPTURE    //Cas Combiné avec tri alphabétique
if Alpha<>nil then
   if Alpha.Checked=True then
   begin ChampOrderBy:='PSA_LIBELLE,';   end;
if (CN1<>nil)and(CN2<>nil)and(CN3<>nil)and(CN4<>nil)and(CN5<>nil) then
  Begin
    if (Cetab.Checked=True) and (Alpha.Checked=True) then ChampOrderBy:='PPU_ETABLISSEMENT,PSA_LIBELLE,';
    if (CN1.Checked=True) and (Alpha.Checked=True)   then ChampOrderBy:='PPU_TRAVAILN1,PSA_LIBELLE,';
    if (CN2.Checked=True) and (Alpha.Checked=True)   then ChampOrderBy:='PPU_TRAVAILN2,PSA_LIBELLE,';
    if (CN3.Checked=True) and (Alpha.Checked=True)   then ChampOrderBy:='PPU_TRAVAILN3,PSA_LIBELLE,';
    if (CN4.Checked=True) and (Alpha.Checked=True)   then ChampOrderBy:='PPU_TRAVAILN4,PSA_LIBELLE,';
    if (CN5.Checked=True) and (Alpha.Checked=True)   then ChampOrderBy:='PPU_CODESTAT,PSA_LIBELLE,';
  End;
if (CL1<>nil)and(CL2<>nil)and(CL3<>nil)and(CL4<>nil) then
  Begin
    if (CL1.Checked=True) and (Alpha.Checked=True) then   ChampOrderBy:='PPU_LIBREPCMB1,PSA_LIBELLE,';
    if (CL2.Checked=True) and (Alpha.Checked=True) then   ChampOrderBy:='PPU_LIBREPCMB2,PSA_LIBELLE,';
    if (CL3.Checked=True) and (Alpha.Checked=True) then   ChampOrderBy:='PPU_LIBREPCMB3,PSA_LIBELLE,';
    if (CL4.Checked=True) and (Alpha.Checked=True) then   ChampOrderBy:='PPU_LIBREPCMB4,PSA_LIBELLE,';
  End;
if TCheckBox(Sender).Name='CALPHA' then
  AffectCritereAlpha(Ecran,TCheckBox(Sender).Checked,'PPU_SALARIE','LIBELLE');
end;

procedure TOF_PGEDITMASSE_ETAT.DateElipsisclick(Sender: TObject);
var key : char;
begin
    key := '*';
    ParamDate (Ecran, Sender, Key);
end;

procedure TOF_PGEDITMASSE_ETAT.ExitEdit(Sender: TObject);
var edit : thedit;
begin
edit:=THEdit(Sender);
if edit <> nil then
  if edit.text<>'' then
    if (VH_Paie.PgTypeNumSal='NUM') and (length(Edit.text)<11) and (isnumeric(edit.text)) then
    edit.text:=AffectDefautCode(edit,10);
end;

procedure TOF_PGEDITMASSE_ETAT.MonnaieInverse(Sender: TObject);
var
ChMonInv : TCheckBox;
begin
ChMonInv := TCheckBox(GetControl('CHMONNAIEINV'));
if (ChMonInv<>nil)  then  PgMonnaieInv:=ChMonInv.checked;
PgTauxConvert:=RendTauxConvertion;
end;



procedure TOF_PGEDITMASSE_ETAT.OnClickBulletin(Sender: TObject);
Var i : integer;
Cumulx: THEDIT;
Cumul : string;
begin
if sender <> nil then
  if Sender is TCHeckBox then
    Begin
    SetControlVisible('GBBULLETIN',TCHeckBox(Sender).Checked);
    SetControltext('DATEDEB',GetControlText('XX_VARIABLEDEB'));
    SetControltext('DATEFIN',GetControlText('XX_VARIABLEFIN'));
    SetControlEnabled('DATEDEB',False);
    SetControlEnabled('DATEFIN',False);
    End;

FOR i:=1 to 6 do
  begin
  Cumulx:=THEDIT(GetControl('XX_VARIABLECUM'+IntToStr(i)));
  if i=1 then  cumul:=VH_PAIE.PGCumul01;
  if i=2 then  cumul:=VH_PAIE.PGCumul02;
  if i=3 then  cumul:=VH_PAIE.PGCumul03;
  if i=4 then  cumul:=VH_PAIE.PGCumul04;
  if i=5 then  cumul:=VH_PAIE.PGCumul05;
  if i=6 then  cumul:=VH_PAIE.PGCumul06;
  IF (Cumul<>'') and (Cumulx<>nil) then
    Cumulx.text:=rechdom('PGCUMULPAIE',Cumul,False);
  end;

For i:=1 to 3 do
 begin
 Cumulx:=THEDIT(GetControl('LBLCUMULMOIS'+IntToStr(i)));
 If (Cumulx=nil) then exit;
 if i=1 then Cumulx.text:=VH_Paie.PgLibCumulMois1;
 if i=2 then Cumulx.text:=VH_Paie.PgLibCumulMois2;
 if i=3 then Cumulx.text:=VH_Paie.PgLibCumulMois3;
 End;

SetControlChecked('PGEDITORG1',VH_Paie.PgEditOrg1);
SetControlChecked('PGEDITORG2',VH_Paie.PgEditOrg2);
SetControlChecked('PGEDITORG3',VH_Paie.PgEditOrg3);
SetControlChecked('PGEDITORG4',VH_Paie.PgEditOrg4);
SetControlChecked('PGEDITCODESTAT',VH_Paie.PgEditCodeStat);
end;

procedure TOF_PGEDITMASSE_ETAT.ForceLibellePaie(Sender: TObject);
begin
SetControlEnabled('DATEDEB',TCheckBox(Sender).Checked);
SetControlEnabled('DATEFIN',TCheckBox(Sender).Checked);
end;

procedure TOF_PGEDITMASSE_ETAT.OnClickChargesSoc(Sender: TObject);
begin
if sender <> nil then
  if Sender is TCHeckBox then
    Begin
    SetControlVisible('GBCHARGESSOC',TCHeckBox(Sender).Checked);
    End;
end;

procedure TOF_PGEDITMASSE_ETAT.OnClickReglement(Sender: TObject);
begin
if sender <> nil then
  if Sender is TCHeckBox then
    Begin
    SetControlVisible('GBREGLEMENT',TCHeckBox(Sender).Checked);
    End;
end;


function TOF_PGEDITMASSE_ETAT.RecupCritereGeneral(Prefixe: String): string;
begin
Result:='';
  Result:=Result+';CALPHA='+GetControlText('CALPHA');
  if GetControlText('CALPHA')='-' then
    Result:=Result+';'+Prefixe+'_SALARIE='+GetControlText('PPU_SALARIE')+';'+Prefixe+'_SALARIE_='+GetControlText('PPU_SALARIE_')
  else
    Result:=Result+';'+Prefixe+'_LIBELLE='+GetControlText('PPU_LIBELLE')+';'+Prefixe+'_LIBELLE_='+GetControlText('PPU_LIBELLE_');
  Result:=Result+';CETAB='+GetControlText('CETAB');
  Result:=Result+';'+Prefixe+'_ETABLISSEMENT='+GetControlText('PPU_ETABLISSEMENT')+';'+Prefixe+'_ETABLISSEMENT_='+GetControlText('PPU_ETABLISSEMENT_');
 if VH_Paie.PGNbreStatOrg>=1 then
    Begin
    Result:=Result+';CN1='+GetControlText('CN1');
    Result:=Result+';'+Prefixe+'_TRAVAILN1='+GetControlText('PPU_TRAVAILN1')+';'+Prefixe+'_TRAVAILN1_='+GetControlText('PPU_TRAVAILN1_');
    END;
  if VH_Paie.PGNbreStatOrg>=2 then
    Begin
    Result:=Result+';CN2='+GetControlText('CN2');
    Result:=Result+';'+Prefixe+'_TRAVAILN2='+GetControlText('PPU_TRAVAILN2')+';'+Prefixe+'_TRAVAILN2_='+GetControlText('PPU_TRAVAILN2_');
    End;
  if VH_Paie.PGNbreStatOrg>=3 then
    Begin
    Result:=Result+';CN3='+GetControlText('CN3');
    Result:=Result+';'+Prefixe+'_TRAVAILN3='+GetControlText('PPU_TRAVAILN3')+';'+Prefixe+'_TRAVAILN3_='+GetControlText('PPU_TRAVAILN3_');
    End;
  if VH_Paie.PGNbreStatOrg>=4 then
    Begin
    Result:=Result+';CN4='+GetControlText('CN4');
    Result:=Result+';'+Prefixe+'_TRAVAILN4='+GetControlText('PPU_TRAVAILN4')+';'+Prefixe+'_TRAVAILN4_='+GetControlText('PPU_TRAVAILN4_');
    End;
  if VH_Paie.PgNbCombo>=1 then
    Begin
    Result:=Result+';CL1='+GetControlText('CL1');
    Result:=Result+';'+Prefixe+'_LIBREPCMB1='+GetControlText('PPU_LIBREPCMB1')+';'+Prefixe+'_LIBREPCMB1_='+GetControlText('PPU_LIBREPCMB1_');
    End;
  if VH_Paie.PgNbCombo>=2 then
    Begin
    Result:=Result+';CL2='+GetControlText('CL2');
    Result:=Result+';'+Prefixe+'_LIBREPCMB2='+GetControlText('PPU_LIBREPCMB2')+';'+Prefixe+'_LIBREPCMB2_='+GetControlText('PPU_LIBREPCMB2_');
    End;
  if VH_Paie.PgNbCombo>=3 then
    Begin
    Result:=Result+';CL3='+GetControlText('CL3');
    Result:=Result+';'+Prefixe+'_LIBREPCMB3='+GetControlText('PPU_LIBREPCMB3')+';'+Prefixe+'_LIBREPCMB3_='+GetControlText('PPU_LIBREPCMB3_');
    End;
  if VH_Paie.PgNbCombo>=4 then
    Begin
    Result:=Result+'CL4=;'+GetControlText('CL4');
    Result:=Result+';'+Prefixe+'_LIBREPCMB4='+GetControlText('PPU_LIBREPCMB4')+';'+Prefixe+'_LIBREPCMB4_='+GetControlText('PPU_LIBREPCMB4_');
    End;
  Result:=Result+';FApercu='+GetControlText('FAPERCU')+';FReduire='+GetControlText('FReduire')+';FCouleur='+GetControlText('FCouleur');

end;



procedure TOF_PGEDITMASSE_ETAT.OnUpdate;
var
StSql,Critere,OrderBy,StDate : String;
DateDebut,DateFin : String;
Temp,Tempo : String;
Pages:TPageControl;
x : integer;
FApercu,FListe,FReduire,okok : Boolean;
begin
inherited;

FApercu:=(GetControlText('FAPERCU')='X');
FListe:=(GetControlText('FLISTE')='X');
FReduire:=(GetControlText('FREDUIRE')='X');
V_PGI.QRCouleur:=(GetControlText('FCOULEUR')='X');

Pages:=TPageControl(GetControl('Pages'));
Temp:=RecupWhereCritere(Pages);
tempo:=''; critere:='';
x:=Pos('(',Temp);
if x>0 then Tempo:=copy(Temp,x,(Length(temp)-5));
if tempo<>'' then critere:='AND '+Tempo;
DateDebut:=''; DateFin:='';
if IsValidDate(GetControlText('XX_VARIABLEDEB')) then
  DateDebut:=USDateTime(StrToDate(GetControlText('XX_VARIABLEDEB')));
if IsValidDate(GetControlText('XX_VARIABLEFIN')) then
  DateFin:=USDateTime(StrToDate(GetControlText('XX_VARIABLEFIN')));

if (DateDebut='') or (DateFin='') then
  Begin
  PGIBox('La session de paie selectionnée est erronée!',Ecran.Caption);
  Exit;
  End;


StDate:=' AND PPU_DATEDEBUT="'+DateDebut+'" AND PPU_DATEFIN="'+DateFin+'"';

//***************** Edition du bulletin de paie *****************
If GetControlText('CKBULLETIN')='X' then
  Begin
  //Critères spécifiques
  PGLanceBulletin:=True;
                     //'+AffectSqlEtat(Ecran)+',
  OrderBy:='ORDER BY '+ChampOrderBy+'PPU_SALARIE,PPU_DATEDEBUT,PPU_DATEFIN,PHB_ORDREETAT,PHB_NATURERUB,PHB_RUBRIQUE' ;
//   StSql:=RendRequeteSQLEtat('BULLETIN','',Critere+StDate,OrderBy);
  okok:=LanceEtat('E','PBP','PBP',FApercu,False,FReduire,Pages,StSql,'',False)>0;
  End;

If GetControlText('CKJOURNALPAIE')='X' then
  Begin
  if Fliste then temp:='PJL' else temp:='PJP';
  StDate:='WHERE PPU_DATEDEBUT="'+DateDebut+'" AND PPU_DATEFIN="'+DateFin+'"';
  OrderBY:='ORDER BY '+ChampOrderBy+' PPU_SALARIE';
  StSql:=RendRequeteSQLEtat('JOURNALPAIE',AjoutChamp,StDate+Critere,OrderBy);
  okok:=LanceEtat('E','PJP',Temp,FApercu,FListe,FReduire,Pages,StSql,'',False);
  End;










end;









Initialization
registerclasses([TOF_PGEDITMASSE_ETAT]);
end.
