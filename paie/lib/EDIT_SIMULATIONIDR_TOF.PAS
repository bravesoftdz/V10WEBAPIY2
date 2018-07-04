{***********UNITE*************************************************
Auteur  ...... : NA
Créé le ...... : 01/09/2006
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : EDIT_SIMULATIONIDR ()
                 Edition de la simulation IDR
Mots clefs ... : TOF;EDIT_SIMULATIONIDR
*****************************************************************}
Unit EDIT_SIMULATIONIDR_TOF ;

Interface

Uses StdCtrls, 
     Controls, 
     Classes, 
{$IFNDEF EAGLCLIENT}
     db, 
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF} 
     mul,
     QRS1,
{$else}
     eMul,
     eQRS1,
     uTob, 
{$ENDIF}
     forms, 
     sysutils, 
     ComCtrls,
     HCtrls, 
     HEnt1, 
     HMsgBox,
     Entpaie,
     p5def,
     Hqry,
     pgoutils2,
     UTOF ; 

Type
  TOF_EDIT_SIMULATIONIDR = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
    private
    procedure change(sender:TObject);
    procedure verifsimul(sender:TObject);
    procedure gestionchampcompl;
    procedure gestionruptchampcompl;
    procedure informeruptchampcompl;
   end ;

Implementation

procedure TOF_EDIT_SIMULATIONIDR.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_EDIT_SIMULATIONIDR.OnDelete ;
begin
  Inherited ;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 04/09/2006
Modifié le ... :   /  /    
Description .. : On update 
Mots clefs ... : 
*****************************************************************}
procedure TOF_EDIT_SIMULATIONIDR.OnUpdate ;
var
  ch_ruptchampcpl : THedit;
  ch1, ch2 : Tcheckbox;
  Pages : Tpagecontrol;
  Sql, Orderby, champcompl, stwhere : string;

begin
  Inherited ;
  Pages := Tpagecontrol(Getcontrol('Pages'));
  stwhere := Recupwherecritere(Pages);
  ch1 := TcheckBox(getcontrol('CPOPUL'));
  ch2 := TcheckBox(getcontrol('CETAB'));
  ch_ruptchampcpl := ThEdit(Getcontrol('XX_RUPTURE3'));

  Champcompl := '';
  if Trim(ch_ruptchampcpl.text) <> '' then champcompl := trim(ch_ruptchampcpl.text);

    orderby := '';
    if (ch1 <> nil)  and (ch1.checked = true) then orderby := 'PSR_POPULATION,' ;
    if (ch2 <> nil)  and (ch2.checked = true) then orderby := orderby + 'PSR_ETABLISSEMENT,';
    if  champcompl <> '' then orderby := orderby + champcompl + ',';

    if orderby <> ''  then orderby := Copy(orderby, 1, length(orderby) - 1);

    if champcompl <> '' then
    Sql := 'select psr_codesimul, psr_salarie, psr_population, psr_etablissement,psr_salaireactu, psr_datedepretr,'+
    ' psr_txcharge,psr_moisacquis, psr_revalorisation, psr_actualisation, psr_probasuivie, psr_probapresence,'+
    ' psr_indemnbrut, psr_nom, psr_prenom,psl_libelle, psl_datecalc, psl_typedepart, psl_methcalcul,'+
    ' psl_tauxactual, '+ Champcompl +' from RESULTSIMUL left join SIMULATION on psr_codesimul = psl_codesimul  '+ stwhere +''
    else
     Sql := 'select psr_codesimul, psr_salarie, psr_population, psr_etablissement,psr_salaireactu, psr_datedepretr,'+
    ' psr_txcharge,psr_moisacquis, psr_revalorisation, psr_actualisation, psr_probasuivie, psr_probapresence,'+
    ' psr_indemnbrut, psr_nom, psr_prenom,psl_libelle, psl_datecalc, psl_typedepart, psl_methcalcul,'+
    ' psl_tauxactual from RESULTSIMUL left join SIMULATION on psr_codesimul = psl_codesimul '+ stwhere +'';

    if orderby <> '' then
    sql := sql + 'order by ' + orderby;

    TFQRS1(Ecran).WhereSQL:=SQL;

end ;

procedure TOF_EDIT_SIMULATIONIDR.OnLoad ;
begin
  Inherited ;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 01/09/2006
Modifié le ... :   /  /    
Description .. : on Argument
Mots clefs ... : 
*****************************************************************}
procedure TOF_EDIT_SIMULATIONIDR.OnArgument (S : String ) ;
var
check : tcheckbox;
i : integer;
codesimul : THValComboBox;
nodossier, sql1,st, pred : string;

begin
  Inherited ;
  nodossier := PgrendNodossier();
  // recherche si population existe dans le dossier
  pred := 'DOS';
  sql1 := 'select PPC_POPULATION,PPO_CODEPOP from ORDREPOPULATION,CODEPOPULATION WHERE PPC_CODEPOP = PPO_CODEPOP'+
  ' and PPO_PREDEFINI = "DOS" and PPO_NODOSSIER = "'+nodossier+'" and PPO_TYPEPOP like "%IDR%"';
  if not existesql(sql1) then
  begin
     // recherche si population existe en standard
     pred := 'STD';
    sql1 := 'select PPC_POPULATION,PPO_CODEPOP from ORDREPOPULATION, CODEPOPULATION WHERE PPC_CODEPOP = PPO_CODEPOP'+
    ' and PPO_PREDEFINI = "STD" and and PPO_TYPEPOP like "%IDR%"';
    if not existesql(sql1) then   pred := 'CEG';

  end;
  st := ' and PPC_PREDEFINI = "'+pred+'"';
  setcontrolproperty('PSR_POPULATION', 'plus', st);


  setcontrolenabled('BVALIDER', false);
  setcontrolvisible('BVALIDER', false);
  Codesimul := THValcomboBox(Getcontrol('PSR_CODESIMUL'));
  if Assigned(codesimul) then codesimul.Onchange := verifsimul;

  check := tcheckbox(getcontrol('CPOPUL'));
  if check <> nil then check.onclick := change;
  check := tcheckbox(getcontrol('CETAB'));
  if check <> nil then check.onclick := change;

  // gestion des champs complémentaires //
  gestionchampcompl;


  setcontrolproperty('TBCOMPLEMENT', 'Tabvisible', (VH_Paie.Pgnbrestatorg > 0) or (VH_Paie.PGLibCodeStat <> ''));
  SetControlProperty('TBCHAMPLIBRE', 'Tabvisible', (VH_Paie.PgNbCombo > 0));

   For I := 1 to 5 do
   Begin
    Check := TCheckBox(GetControl('CN'+IntToStr(I)));
    If Check <> nil Then Check.OnClick := Change;
   End;
   For I := 1 to 4 do
   Begin
    Check := TCheckBox(GetControl('CL'+IntToStr(I)));
    If Check <> nil Then Check.OnClick := Change;
   End;

end ;

procedure TOF_EDIT_SIMULATIONIDR.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_EDIT_SIMULATIONIDR.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_EDIT_SIMULATIONIDR.OnCancel () ;
begin
  Inherited ;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 05/09/2006
Modifié le ... :   /  /    
Description .. : Vérification de la saisie du code simulation
Mots clefs ... : 
*****************************************************************}
procedure TOF_EDIT_SIMULATIONIDR.verifsimul(sender : Tobject);
var
codesimul , sql1, popul, listewhere,st,selectpop : string;
qq: Tquery;

begin

codesimul := THValcomboBox(getcontrol('PSR_CODESIMUL')).value;
  if codesimul <> '' then
  begin
    setcontrolenabled('BVALIDER', true);
    setcontrolvisible('BVALIDER', true);
    // Recherche des populations traités dans la simulation
    sql1 := 'select PSL_SELECTPOP FROM SIMULATION WHERE PSL_CODESIMUL = "'+codesimul+'"';
    QQ:= opensql(sql1,true);
    if not QQ.EOF then
    begin
      selectpop := QQ.findfield('PSL_SELECTPOP').asstring;
      popul := READTOKENST(selectpop);
      while popul <> '' do                                // lecture de chaque population
      begin
      listewhere :=  listewhere + '"'+popul+'" ,' ;
      popul := READTOKENST(selectpop);
      end;

      if listewhere <> '' then
      begin                // si au moins une population
      listewhere := '(' + copy(listewhere, 1, length(listewhere) - 1) + ')';

      St := 'AND PPC_POPULATION in  '+listewhere+'';
      setcontrolproperty('PSR_POPULATION', 'plus', st);
      end;
    end;
  end;

end;


{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 01/09/2006
Modifié le ... :   /  /    
Description .. : Saisie d'une rupture
Mots clefs ... : 
*****************************************************************}
procedure TOF_EDIT_SIMULATIONIDR.change(sender : Tobject);
var
ch_ruptpopul, ch_ruptetab : Thedit;
checketab, checkpopul : Tcheckbox;

begin
ch_ruptpopul := THedit(getcontrol('XX_RUPTURE1'));
ch_ruptetab := THedit(getcontrol('XX_RUPTURE2'));
checketab := Tcheckbox(Getcontrol('CETAB'));
checkpopul := Tcheckbox(Getcontrol('CPOPUL'));

// Gestion des ruptures pour les champs complémentaires (rend visible ou non le champ rupture)
Gestionruptchampcompl;

// Recherche des ruptures des champs complémentaires et alimente les champs XX_VARIABLE et XX_RUPTURE3//
informeruptchampcompl;


if ch_ruptpopul <> nil then
begin
 ch_ruptpopul.text := '';
 if (checkpopul <> nil) and (checkpopul.checked = true) then ch_ruptpopul.text := 'PSR_POPULATION';
end;

if ch_ruptetab <> nil then
begin
 ch_ruptetab.text := '';
 if (checketab <> nil) and (checketab.checked = true) then ch_ruptetab.text := 'PSR_ETABLISSEMENT';
end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 01/09/2006
Modifié le ... :   /  /
Description .. : Gestion des champs complémentaires
Mots clefs ... : 
*****************************************************************}
procedure TOF_EDIT_SIMULATIONIDR.gestionchampcompl;
var
  ch1, ch2 : string;
  I: Integer;
  Champ: THValComboBox;
  Check: TCheckBox;

Begin
  VisibiliteStat(GetControl('PSR_CODESTAT'), GetControl('TPSR_CODESTAT'));
  Champ := THValComboBox(GetControl('PSR_CODESTAT'));
  If (Champ <> nil) And (Champ.Visible = True) Then
  Begin
     Check := TCheckBox(GetControl('CN5'));
     If Check <> nil Then
     Begin
       Check.Visible := True;
       Check.Enabled := True;
     End;
  End;

  For I := 1 to 4 do
  Begin
    Ch1 := 'PSR' + '_TRAVAILN' + IntToStr(I);
    Ch2 := 'TPSR' + '_TRAVAILN' + IntToStr(I);
    VisibiliteChampSalarie(IntToStr(I), GetControl(Ch1), GetControl(Ch2));
    Champ := THValComboBox(GetControl(Ch1));
    If (Champ <> nil) And (Champ.Visible = True) Then
    Begin
       Check := TCheckBox(GetControl('CN'+IntToStr(I)));
       If Check <> nil Then
       Begin
         Check.Visible := True;
         Check.Enabled := True;
       End;
    End;
  End;

  For I := 1 to 4 do
  Begin
    Ch1 := 'PSR' + '_LIBREPCMB' + IntToStr(I);
    Ch2 := 'TPSR' +  '_LIBREPCMB' + IntToStr(I);
    VisibiliteChampLibreSal(IntToStr(I), GetControl(Ch1), GetControl(Ch2));
    Champ := THValComboBox(GetControl(Ch1));
    If (Champ <> nil) And (Champ.Visible = True) Then
    Begin
       Check := TCheckBox(GetControl('CL'+IntToStr(I)));
       If Check <> nil Then
       Begin
         Check.Visible := True;
         Check.Enabled := True;
       End;
    End;
  End;
End;

{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 01/09/2006
Modifié le ... :   /  /    
Description .. : Gestion des ruptures des champs complémentaires (pour 
Suite ........ : avoir une seule rupture)
Mots clefs ... : 
*****************************************************************}
Procedure TOF_EDIT_SIMULATIONIDR.gestionruptchampcompl;
var
  TabLieuTravail : array[1..10] of TCheckBox;
  PosCheck,PosUnCheck,i : integer;
  Ok : boolean ;

Begin
  TabLieuTravail[1]:=TCheckBox(GetControl('CN1'));
  TabLieuTravail[2]:=TCheckBox(GetControl('CN2'));
  TabLieuTravail[3]:=TCheckBox(GetControl('CN3'));
  TabLieuTravail[4]:=TCheckBox(GetControl('CN4'));
  TabLieuTravail[5]:=TCheckBox(GetControl('CN5'));
  TabLieuTravail[6]:=TCheckBox(GetControl('CL1'));
  TabLieuTravail[7]:=TCheckBox(GetControl('CL2'));
  TabLieuTravail[8]:=TCheckBox(GetControl('CL3'));
  TabLieuTravail[9]:=TCheckBox(GetControl('CL4'));
  TabLieuTravail[10]:=nil;
  PosUnCheck:=0;
  PosCheck:=0;

  For i:=1 to 9 do
     If (TabLieuTravail[i]<>nil) Then Ok:=False Else Begin Ok:=True; break; End;

  If Ok=False Then
  Begin
    //Coche une rupture
    For i:=1 to 9 do
      If (TabLieuTravail[i].checked=True) Then PosCheck:=i;
    If PosCheck > 0 Then
      For i:=1 to 9 do
         If i<>PosCheck then TabLieuTravail[i].enabled:=False;

    //Décoche une rupture ,  rend enable(True) les autres champs de rupture
    For i:=1 to 9 do
      If (TabLieuTravail[i].checked=False) and (TabLieuTravail[i].enabled=True) then PosUnCheck:=i;
    If (PosCheck=0) and (PosUnCheck>0) then
      For i:=1 to 9 do
         TabLieuTravail[i].enabled:=True;
  End;
End;

{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 01/09/2006
Modifié le ... :   /  /    
Description .. : Alimente les champs VARIABLE et RUPTURE3 avec le 
Suite ........ : nom du champ coché en rupture
Mots clefs ... : 
*****************************************************************}
Procedure TOF_EDIT_SIMULATIONIDR.InformeRuptchampcompl;
var
  CN1,CN2,CN3,CN4,CN5,CL1,CL2,CL3,CL4:TCheckBox;
  Rupture,Champ1 : THEdit;

Begin
  CN1:=TCheckBox(GetControl('CN1'));
  CN2:=TCheckBox(GetControl('CN2'));
  CN3:=TCheckBox(GetControl('CN3'));
  CN4:=TCheckBox(GetControl('CN4'));
  CN5:=TCheckBox(GetControl('CN5'));
  CL1:=TCheckBox(GetControl('CL1'));
  CL2:=TCheckBox(GetControl('CL2'));
  CL3:=TCheckBox(GetControl('CL3'));
  CL4:=TCheckBox(GetControl('CL4'));
  Rupture:=THEdit(GetControl('XX_RUPTURE3'));
  Champ1 :=THEdit(GetControl('XX_VARIABLE'));

  If (Champ1<>nil) and (Rupture<>nil) Then
      Champ1.text:='';Rupture.Text:='';
  If (CN1<>nil) and (CN2<>nil) and (CN3<>nil) and (CN4<>nil) and (CN5<>nil) Then
     If (Champ1<>nil) and (Rupture<>nil) Then
     Begin
       If (CN1.Checked=True) Then
       Begin
          Champ1.text:=VH_Paie.PGLibelleOrgStat1;
          Rupture.Text:='PSR_TRAVAILN1';
       End;
       If (CN2.Checked=True) Then
       Begin
          Champ1.text:=VH_Paie.PGLibelleOrgStat2;
          Rupture.Text:= 'PSR_TRAVAILN2';
       End;
       If (CN3.Checked=True) then
       Begin
          Champ1.text:=VH_Paie.PGLibelleOrgStat3;
          Rupture.Text:= 'PSR_TRAVAILN3';
       End;
       If (CN4.Checked=True) then
       Begin
          Champ1.text:=VH_Paie.PGLibelleOrgStat4;
          Rupture.Text:= 'PSR_TRAVAILN4';
       End;
       If (CN5.Checked=True) then
       Begin
          Champ1.text:=VH_Paie.PGLibCodeStat;
          Rupture.Text:= 'PSR_CODESTAT';
       End;
     End;
  If (CL1<>nil) and (CL2<>nil) and (CL3<>nil) and (CL4<>nil) Then
     If (Champ1<>nil) and (Rupture<>nil) Then
     Begin
       If (CL1.Checked=True) Then
       Begin
          Champ1.text:=VH_Paie.PgLibCombo1;
          Rupture.Text:= 'PSR_LIBREPCMB1';
       End;
       If (CL2.Checked=True) Then
       Begin
          Champ1.text:=VH_Paie.PgLibCombo2;
          Rupture.Text:= 'PSR_LIBREPCMB2';
       End;
       If (CL3.Checked=True) Then
       Begin
          Champ1.text:=VH_Paie.PgLibCombo3;
          Rupture.Text:= 'PSR_LIBREPCMB3';
       End;
       If (CL4.Checked=True) Then
       Begin
          Champ1.text:=VH_Paie.PgLibCombo4;
          Rupture.Text:= 'PSR_LIBREPCMB4';
       End;
     End;
End;

Initialization
  registerclasses ( [ TOF_EDIT_SIMULATIONIDR ] ) ; 
end.
