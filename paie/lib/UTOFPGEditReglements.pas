{***********UNITE*************************************************
Auteur  ...... : PAIE PGI
Créé le ...... : 31/07/2001
Modifié le ... :   /  /
Description .. : Edition des réglements
Mots clefs ... : PAIE;REGLEMENT
*****************************************************************
PT1   : 31/07/2001 SB V547 Modification champ suffixe MODEREGLE
PT2   : 11/09/2001 SB V547 Champ PPU_PGMODEREGLE DataType non renseigné
PT3   : 29/01/2002 SB V571 Sélection alphanumérique du code salarié
PT4   : 11/03/2002 SB V571 Ajout champ dossier
PT5   : 18/04/2002 SB V571 Fiche de bug n°10087 : V_PGI_env.LibDossier non
                           renseigné en Mono
PT6   : 20/11/2002 SB V591 Traitements Etats chainés
PT7   : 15/01/2003 SB V591 FQ 10448 Dysfonctionnement tri assemblé
PT8   : 14/02/2003 SB V595 FQ 10448 Si pls date paie dans session alors
                           dysfonctionnement Tri..
                           Les enr. doivent être trié par mensualité de paie
PT9   : 23/07/2003 SB V_42 FQ 10586 Edition acompte : ajout critère
PT10  : 06/10/2003 SB V_42 Edition Acompte : Suppression tri pas PSD_DATEDEBUT,
                           R_DOMICILIATION
PT11  : 29/10/2003 SB V_42 FQ 10586
PT12  : 10/08/2004 VG V_50 Ajout d'un critère "Exclure les salariés sortis"
                           FQ N°11275
PT13  : 19/05/2006 SB V_65 Appel des fonctions communes pour identifier V_PGI.driver
PT14  : 16/01/07 V8_00 FCO Mise en place filtrage des habilitations/poupulations
PT15  : 14/05/07 V7_xx RMA Ajout volets compléments et champ_libre pour l'impression des
                           mode de réglements FQ 13961
}
unit UTOFPGEditReglements;

interface
uses StdCtrls, Controls, Classes, sysutils,
  {$IFDEF EAGLCLIENT}
  eQRS1,
  {$ELSE}
   QRS1,
  {$ENDIF}
  HCtrls, HEnt1,  UTOF,
  ComCtrls, HQry,  //PT15
  ParamSoc, UTOFPGEtats;


type
  TOF_PGMODEREGLE_ETAT = class(TOF_PGEtats)
    procedure OnArgument(Arguments: string); override;
    procedure OnLoad; override;
    Procedure OnUpdate ; override ; //PT15
    procedure MonnaieInverse(Sender: Tobject);
    procedure Change(Sender: Tobject);
    procedure ExitEdit(Sender: TObject);
  private
    DD, DF: TDateTime; //PT6
    Origine: string; //PT6
    Procedure ControleChampCompl(PrefTable : String) ; //PT15
    Procedure ControleCheckRupture ;                   //PT15
    Procedure InformeRupture(PrefTable : String);      //PT15
  end;

  TOF_PGRIB_ETAT = class(TOF)
    procedure OnArgument(Arguments: string); override;
    procedure ExitEdit(Sender: TObject);
    procedure OnClickSalarieSortie(Sender: TObject);
    procedure OnUpdate ; override;


  end;

  TOF_PGACOMPTE_ETAT = class(TOF)
    procedure OnArgument(Arguments: string); override;
    procedure ChangeModeReglement(Sender: TObject);
    procedure ExitEdit(Sender: TObject);
    procedure OnClickCritere(Sender: Tobject);
  end;

implementation

uses PgEditOutils, PGEditOutils2,EntPaie, PgOutils2, PGCommun,
  P5Def; //PT14

procedure TOF_PGMODEREGLE_ETAT.OnArgument(Arguments: string);
var
  CDd, CDf: THEdit;
  CEtab, CAlpha, CMont, ChMonInv: TCheckBox;
  Defaut, Convert: THEDIT;
  Combo: THValComboBox;
  Ok: Boolean;
  Min, Max, DebPer, FinPer, ExerPerEncours, Origine: string;
  Ckeuro: TCheckBox;
  I : Integer;        //PT15
  Check : TCheckBox;  //PT15
begin
  inherited;
  //DEB PT6
  DF := idate1900;
  if (Arguments <> '') then
  begin
    if (Pos('CHAINES', Arguments) > 0) then
    begin
      Origine := ReadTokenSt(Arguments);
      if trim(Arguments) <> '' then
      begin
        DD := StrToDate(ReadTokenSt(Arguments));
        DF := StrToDate(ReadTokenSt(Arguments));
      end;
    end;
  end;
  //FIN PT6
  SetControlProperty('PPU_PGMODEREGLE', 'Datatype', 'PGMODEREGLE'); //PT2
  Ckeuro := TCheckBox(GetControl('CKEURO'));
  if Ckeuro <> nil then Ckeuro.Checked := VH_Paie.PGTenueEuro;
  Defaut := ThEdit(getcontrol('DOSSIER'));
  if Defaut <> nil then
    //   Defaut.text:=V_PGI_env.LibDossier; //PT5 Mise en commentaire
    Defaut.text := GetParamSoc('SO_LIBELLE');

  //Affectation des Valeurs par défaut
  RecupMinMaxTablette('PG', 'SALARIES', 'PSA_SALARIE', Min, Max);
  Defaut := ThEdit(getcontrol('PPU_SALARIE'));
  if Defaut <> nil then
  begin
    Defaut.text := Min;
    Defaut.OnExit := ExitEdit;
  end;
  Defaut := ThEdit(getcontrol('PPU_SALARIE_'));
  if Defaut <> nil then
  begin
    Defaut.text := Max;
    Defaut.OnExit := ExitEdit;
  end;
  RecupMinMaxTablette('PG', 'ETABLISS', 'ET_ETABLISSEMENT', Min, Max);
  Defaut := ThEdit(getcontrol('PPU_ETABLISSEMENT'));
  if Defaut <> nil then Defaut.text := Min;
  Defaut := ThEdit(getcontrol('PPU_ETABLISSEMENT_'));
  if Defaut <> nil then Defaut.text := Max;
  Combo := ThValComboBox(getcontrol('PPU_PGMODEREGLE')); {PT1}
  if Combo <> nil then Combo.ItemIndex := 0;

  Convert := ThEdit(getcontrol('XX_VARIABLECONV'));
  if convert <> nil then Convert.text := '1';

  ChMonInv := TCheckBox(GetControl('CHMONNAIEINV'));
  if ChMonInv <> nil then ChMonInv.OnClick := MonnaieInverse;

  CDd := THEdit(GetControl('PPU_DATEDEBUT'));
  CDf := THEdit(GetControl('PPU_DATEFIN'));
  ok := RendPeriodeEnCours(ExerPerEncours, DebPer, FinPer);
  if Ok = True then
  begin
    if CDd <> nil then CDd.text := DebPer;
    if CDf <> nil then CDf.text := FinPer;
  end;

  //Evenement sur Onclick Rupture Sos Total
  CEtab := TCheckBox(GetControl('CETAB'));
  if Cetab <> nil then Cetab.OnClick := Change;
  CAlpha := TCheckBox(GetControl('CALPHA'));
  if CAlpha <> nil then CAlpha.OnClick := Change;
  CMont := TCheckBox(GetControl('CMONT'));
  if CMont <> nil then CMont.OnClick := Change;

  //DEB PT6 Affect critère standard
  if (origine = 'CHAINES') and (DF <> idate1900) then
  begin
    if CDd <> nil then CDd.text := DateToStr(DD);
    if CDf <> nil then CDf.text := DateToStr(DF);
  end;
  //FIN PT6
  //PT15 Debut Ajout ==>
  ControleChampCompl ('PPU');

  SetControlProperty('TBCOMPLEMENT', 'Tabvisible', (VH_Paie.PGNbreStatOrg > 0) or (VH_Paie.PGLibCodeStat <> ''));
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
  //PT15 Fin Ajout <====
end;

procedure TOF_PGMODEREGLE_ETAT.OnLoad;
begin
  inherited;
  //DEB PT6 Affect critère standard
  if (origine = 'CHAINES') and (DF <> idate1900) then
  begin
    SetControlText('PPU_DATEDEBUT', DateToStr(DD));
    SetControlText('PPU_DATEFIN', DateToStr(DF));
  end;
  //FIN PT6
end;

procedure TOF_PGMODEREGLE_ETAT.Change(Sender: Tobject);
var
  CEtab, CAlpha, CMont: TCheckBox;
  rupture, xorder: THEdit;

begin
  //PT15 Debut Ajout ==>
  ControleCheckRupture;
  InformeRupture('PPU');
  //PT15 Fin Ajout <====
  CEtab := TCheckBox(GetControl('CETAB'));
  CAlpha := TCheckBox(GetControl('CALPHA'));
  CMont := TCheckBox(GetControl('CMONT'));
  rupture := THEdit(GetControl('XX_RUPTURE1'));
  xorder := THEdit(GetControl('XX_ORDERBY'));

  if (Cetab <> nil) and (CAlpha <> nil) and (CMont <> nil) and (rupture <> nil) and (xorder <> nil) then
  begin
    rupture.text := ''; //PT7 Ajout
    xorder.text := ''; //PT7 Ajout

    if Cetab.checked = True then
    begin
      rupture.text := 'PPU_ETABLISSEMENT'; //PPU_ETABLISSEMENT
      { DEB PT8 }
      if (PGisOracle) then { PT13 }
        xorder.text := ' To_Char (PPU_DATEFIN,"YY"),To_Char (PPU_DATEFIN,"MM"),PPU_PGMODEREGLE,PPU_SALARIE ' {PT1}
      else
        if (PGisMssql) or (PGisSYBASE) then  { PT13 }
        xorder.text := ' DATEPART(YEAR,PPU_DATEFIN),DATEPART(MONTH,PPU_DATEFIN),PPU_PGMODEREGLE,PPU_SALARIE ' {PT1}
      else
        xorder.text := ' YEAR(PPU_DATEFIN),MONTH(PPU_DATEFIN),PPU_PGMODEREGLE,PPU_SALARIE '; {PT1}
      { END PT8 }
    end;
    if CAlpha.checked = True then
    begin
      CMont.Checked := False;
      CMont.enabled := False; {PT7 Mise en commentaire rupture.text:='';}
      { DEB PT8 }
       if (PGisOracle) then  { PT13 }
        xorder.text := ' To_Char (PPU_DATEFIN,"YY"),To_Char (PPU_DATEFIN,"MM"),PPU_PGMODEREGLE,PPU_LIBELLE ' {PT1}
      else
        if (PGisMssql) or (PGisSYBASE) then  { PT13 }
        xorder.text := ' DATEPART(YEAR,PPU_DATEFIN),DATEPART(MONTH,PPU_DATEFIN),PPU_PGMODEREGLE,PPU_LIBELLE ' {PT1}
      else
        xorder.text := ' YEAR(PPU_DATEFIN),MONTH(PPU_DATEFIN),PPU_PGMODEREGLE,PPU_LIBELLE '; {PT1}
      { FIN PT8 }
    end else CMont.enabled := True;
    if CMont.checked = True then
    begin
      CAlpha.Checked := False;
      CAlpha.enabled := False; {PT7 Mise en commentaire rupture.text:='';  }
      { DEB PT8 }
      if (PGisOracle) then  { PT13 }
        xorder.text := ' To_Char (PPU_DATEFIN,"YY"),To_Char (PPU_DATEFIN,"MM"),PPU_PGMODEREGLE,PPU_CNETAPAYER DESC ' {PT1}
      else
        if (PGisMssql) or (PGisSYBASE) then { PT13 }
        xorder.text := ' DATEPART(YEAR,PPU_DATEFIN),DATEPART(MONTH,PPU_DATEFIN),PPU_PGMODEREGLE,PPU_CNETAPAYER DESC ' {PT1}
      else
        xorder.text := ' YEAR(PPU_DATEFIN),MONTH(PPU_DATEFIN),PPU_PGMODEREGLE,PPU_CNETAPAYER DESC '; {PT1}
      { FIN PT8 }
    end else CAlpha.enabled := True;
    {PT7 Mise en commentaire
    if (Cetab.checked=False) and  (CAlpha.checked=False) and (CMont.checked=False) then
      begin
      rupture.text:=''; xorder.text:=''; end;    }
  end;
  if TCheckBox(Sender).Name = 'CALPHA' then //PT3
    AffectCritereAlpha(Ecran, TCheckBox(Sender).Checked, 'PPU_SALARIE', 'PPU_LIBELLE');

end;

procedure TOF_PGMODEREGLE_ETAT.MonnaieInverse(Sender: TObject);
var
  ChMonInv: TCheckBox;
  Convert: ThEdit;
begin
  ChMonInv := TCheckBox(GetControl('CHMONNAIEINV'));
  Convert := ThEdit(getcontrol('XX_VARIABLECONV'));
  if (ChMonInv <> nil) and (convert <> nil) then
  begin
    if ChMonInv.checked = False then Convert.text := '1';
    if ChMonInv.checked = True then Convert.text := RendTauxConvertion;
  end;
end;

procedure TOF_PGMODEREGLE_ETAT.ExitEdit(Sender: TObject);
var
  edit: thedit;
begin
  edit := THEdit(Sender);
  if edit <> nil then
    if edit.text <> '' then
      if (VH_Paie.PgTypeNumSal = 'NUM') and (length(Edit.text) < 11) and (isnumeric(edit.text)) then
        edit.text := AffectDefautCode(edit, 10);
end;

//PT15 Debut Ajout ==>
Procedure TOF_PGMODEREGLE_ETAT.ControleChampCompl (PrefTable : String) ;
var
  Ch1, Ch2, Ch3: string;
  I: Integer;
  Champ: THValComboBox;
  Check: TCheckBox;

Begin
  If Trim(PrefTable) = '' Then Exit;
  Ch1 := PrefTable + '_CODESTAT';
  Ch2 := 'T' + PrefTable + '_CODESTAT';
  Ch3 := 'R_CODESTAT';
  VisibiliteStat(GetControl(Ch1), GetControl(Ch2), GetControl(Ch3));
  VisibiliteStat(GetControl(Ch1 + '_'), GetControl(Ch2 + '_'));
  VisibiliteStat(GetControl(Ch1 + '__'), GetControl(Ch2 + '__'));
  Champ := THValComboBox(GetControl(Ch1));
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
    Ch1 := PrefTable + '_TRAVAILN' + IntToStr(I);
    Ch2 := 'T' + PrefTable + '_TRAVAILN' + IntToStr(I);
    Ch3 := 'R_TRAVAILN' + IntToStr(I);
    VisibiliteChampSalarie(IntToStr(I), GetControl(Ch1), GetControl(Ch2), GetControl(Ch3));
    VisibiliteChampSalarie(IntToStr(I), GetControl(Ch1 + '_'), GetControl(Ch2 + '_'));
    VisibiliteChampSalarie(IntToStr(I), GetControl(Ch1 + '__'), GetControl(Ch2 + '__'));
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
    Ch1 := PrefTable + '_LIBREPCMB' + IntToStr(I);
    Ch2 := 'T' + PrefTable + '_LIBREPCMB' + IntToStr(I);
    Ch3 := 'R_LIBREPCMB' + IntToStr(I);
    VisibiliteChampLibreSal(IntToStr(I), GetControl(Ch1), GetControl(Ch2), GetControl(Ch3));
    VisibiliteChampLibreSal(IntToStr(I), GetControl(Ch1 + '_'), GetControl(Ch2 + '_'));
    VisibiliteChampLibreSal(IntToStr(I), GetControl(Ch1 + '__'), GetControl(Ch2 + '__'));
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

Procedure TOF_PGMODEREGLE_ETAT.ControleCheckRupture ;
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

Procedure TOF_PGMODEREGLE_ETAT.InformeRupture(PrefTable : String);
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
  Rupture:=THEdit(GetControl('XX_RUPTURE2'));
  Champ1 :=THEdit(GetControl('XX_VARIABLE1'));

  If (Champ1<>nil) and (Rupture<>nil) Then
      Champ1.text:='';Rupture.Text:='';
  If (CN1<>nil) and (CN2<>nil) and (CN3<>nil) and (CN4<>nil) and (CN5<>nil) Then
     If (Champ1<>nil) and (Rupture<>nil) Then
     Begin
       If (CN1.Checked=True) Then
       Begin
          Champ1.text:=VH_Paie.PGLibelleOrgStat1;
          Rupture.Text:=PrefTable+'_TRAVAILN1';
       End;
       If (CN2.Checked=True) Then
       Begin
          Champ1.text:=VH_Paie.PGLibelleOrgStat2;
          Rupture.Text:=PrefTable+'_TRAVAILN2';
       End;
       If (CN3.Checked=True) then
       Begin
          Champ1.text:=VH_Paie.PGLibelleOrgStat3;
          Rupture.Text:=PrefTable+'_TRAVAILN3';
       End;
       If (CN4.Checked=True) then
       Begin
          Champ1.text:=VH_Paie.PGLibelleOrgStat4;
          Rupture.Text:=PrefTable+'_TRAVAILN4';
       End;
       If (CN5.Checked=True) then
       Begin
          Champ1.text:=VH_Paie.PGLibCodeStat;
          Rupture.Text:=PrefTable+'_CODESTAT';
       End;
     End;
  If (CL1<>nil) and (CL2<>nil) and (CL3<>nil) and (CL4<>nil) Then
     If (Champ1<>nil) and (Rupture<>nil) Then
     Begin
       If (CL1.Checked=True) Then
       Begin
          Champ1.text:=VH_Paie.PgLibCombo1;
          Rupture.Text:=PrefTable+'_LIBREPCMB1';
       End;
       If (CL2.Checked=True) Then
       Begin
          Champ1.text:=VH_Paie.PgLibCombo2;
          Rupture.Text:=PrefTable+'_LIBREPCMB2';
       End;
       If (CL3.Checked=True) Then
       Begin
          Champ1.text:=VH_Paie.PgLibCombo3;
          Rupture.Text:=PrefTable+'_LIBREPCMB3';
       End;
       If (CL4.Checked=True) Then
       Begin
          Champ1.text:=VH_Paie.PgLibCombo4;
          Rupture.Text:=PrefTable+'_LIBREPCMB4';
       End;
     End;
End;

Procedure TOF_PGMODEREGLE_ETAT.OnUpdate ;
var
  Pages :TPageControl;
  rupture, xorder, rupture2 : String;
  SQL, Where ,AjoutSQL ,AjoutWhere ,AjoutOrder :String;
  P : Integer;

Begin
  Inherited ;
  Pages := TPageControl(GetControl('Pages'));
  Where := RecupWhereCritere(Pages);

  P := Pos('ORDER',Where);
  If P > 0 Then Where := Copy(Where,1,P-2);
  rupture  := Trim(GetControlText('XX_RUPTURE1'));
  rupture2 := Trim(GetControlText('XX_RUPTURE2'));
  xorder   := Trim(GetControlText('XX_ORDERBY'));

  If rupture2 <>'' Then
  Begin
     SQL := 'SELECT PPU_ETABLISSEMENT,PPU_SALARIE,PPU_DATEDEBUT,PPU_DATEFIN,' +
            'PPU_PGMODEREGLE,PPU_PAYELE,PPU_CNETAPAYER,PPU_AUXILIAIRE,PPU_RIBSALAIRE,' +
            'PPU_BANQUEEMIS,PPU_ECHEANCE,PPU_TOPREGLE,' + rupture2 + ',' +
            'R_AUXILIAIRE,R_ETABBQ,R_GUICHET,R_NUMEROCOMPTE,R_CLERIB,R_DOMICILIATION ' +
            'FROM PAIEENCOURS ' +
            'LEFT JOIN RIB ON R_AUXILIAIRE=PPU_AUXILIAIRE AND R_SALAIRE="X" ';
     SQL := SQL + Where;
     SQL := SQL + ' ORDER BY ';
     If rupture <> '' Then SQL := SQL + rupture + ',' + rupture2 + ','
                      Else SQL := SQL + rupture2 + ',';
     If xorder <> '' Then SQL := SQL + xorder
                     Else SQL := SQL + 'PPU_DATEFIN,PPU_PGMODEREGLE DESC,PPU_SALARIE,PPU_LIBELLE';

     TFQRS1(Ecran).WhereSQL:=SQL;
  End;
End;
//PT15 Fin Ajout <====

{ TOF_PGRIB_ETAT }
{-------------------------------------------------------------------------------
                                 RIB
-------------------------------------------------------------------------------}
procedure TOF_PGRIB_ETAT.OnArgument(Arguments: string);
var
  Defaut: THEDIT;
  Min, Max: string;
  Check : TCheckBox;
begin
  inherited;
  //Affectation des Valeurs par défaut
  RecupMinMaxTablette('PG', 'SALARIES', 'PSA_SALARIE', Min, Max);
  Defaut := ThEdit(getcontrol('PSA_SALARIE'));
  if Defaut <> nil then
  begin
    Defaut.text := Min;
    Defaut.OnExit := ExitEdit;
  end;
  Defaut := ThEdit(getcontrol('PSA_SALARIE_'));
  if Defaut <> nil then
  begin
    Defaut.text := Max;
    Defaut.OnExit := ExitEdit;
  end;
  RecupMinMaxTablette('PG', 'ETABLISS', 'ET_ETABLISSEMENT', Min, Max);
  Defaut := ThEdit(getcontrol('PSA_ETABLISSEMENT'));
  if Defaut <> nil then Defaut.text := Min;
  Defaut := ThEdit(getcontrol('PSA_ETABLISSEMENT_'));
  if Defaut <> nil then Defaut.text := Max;
  //SetControltext('DOSSIER',V_PGI_env.LibDossier);   //DEb PT4  //PT5 Mise en commentaire
  SetControlText('DOSSIER', GetParamSoc('SO_LIBELLE')); //Fin PT4

//PT12
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
//FIN PT12
end;

procedure TOF_PGRIB_ETAT.ExitEdit(Sender: TObject);
var
  edit: thedit;
begin
  edit := THEdit(Sender);
  if edit <> nil then
    if edit.text <> '' then
      if (VH_Paie.PgTypeNumSal = 'NUM') and (length(Edit.text) < 11) and (isnumeric(edit.text)) then
        edit.text := AffectDefautCode(edit, 10);

end;

//PT12
{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 09/08/2004
Modifié le ... :   /  /
Description .. : Click sur la check-Box "Exclure les salariés sortis"
Mots clefs ... : PAIE;REGLEMENT
*****************************************************************}
procedure TOF_PGRIB_ETAT.OnClickSalarieSortie(Sender: TObject);
begin
SetControlenabled('DATEARRET',(GetControltext('CKSORTIE')='X'));
SetControlenabled('TDATEARRET',(GetControltext('CKSORTIE')='X'));
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 10/08/2004
Modifié le ... :   /  /
Description .. : Fonction OnLoad pour pouvoir modifier les critères en
Suite ........ : fonction de la check-box "Exclure les salariés sortis"
Mots clefs ... : PAIE;REGLEMENT
*****************************************************************}
procedure TOF_PGRIB_ETAT.OnUpdate ;
var
CWhere, Etab1, Etab2, OrderBy, requete, Sal1, Sal2, SQL, StDateArret : string;
DateArret : TdateTime;
begin
Inherited ;
Etab1:=GetControlText('PSA_ETABLISSEMENT');
Etab2:=GetControlText('PSA_ETABLISSEMENT_');
Sal1:=GetControlText('PSA_SALARIE');
Sal2:=GetControlText('PSA_SALARIE_');

requete:= 'SELECT PSA_SALARIE, PSA_LIBELLE, PSA_PRENOM, PSA_ETABLISSEMENT,'+
          ' RIB.*,ET_LIBELLE  FROM SALARIES'+
          ' LEFT JOIN RIB ON'+
          ' PSA_AUXILIAIRE=R_AUXILIAIRE'+
          ' LEFT JOIN ETABLISS ON'+
          ' ET_ETABLISSEMENT=PSA_ETABLISSEMENT';
CWhere:= ' WHERE R_AUXILIAIRE<>""';
OrderBy:= ' ORDER BY PSA_ETABLISSEMENT,PSA_SALARIE';

if ((Sal1<>'') or (Sal2<>'') or (Etab1<>'') or (Etab2<>'')) then
   begin
   CWhere:= CWhere+' AND (';

   If (Sal1<>'') Then
      CWhere:= CWhere+' PSA_SALARIE>="'+Sal1+'"';

   If (Sal2<>'') Then
      begin
      If (Sal1<>'') Then
         CWhere:= CWhere+' AND';
      CWhere:= CWhere+' PSA_SALARIE<="'+Sal2+'"';
      end;

   If (Etab1<>'') Then
      begin
      If ((Sal1<>'') or (Sal2<>'')) Then
         CWhere:= CWhere+' AND';
      CWhere:= CWhere+' PSA_ETABLISSEMENT>="'+Etab1+'"';
      end;

   If (Etab2<>'') Then
      begin
      If ((Sal1<>'') or (Sal2<>'') or (Etab1<>'')) Then
         CWhere:= CWhere+' AND';
      CWhere:= CWhere+' PSA_ETABLISSEMENT<="'+Etab2+'"';
      end;
   CWhere:= CWhere+')';
   end;

if TCheckBox(GetControl('CKSORTIE'))<>nil then
   Begin
   if (GetControlText('CKSORTIE')='X') and
      (IsValidDate(GetControlText('DATEARRET')))then
      Begin
      DateArret:= StrtoDate(GetControlText('DATEARRET'));
      StDateArret:= ' AND (PSA_DATESORTIE>="'+UsDateTime(DateArret)+'" OR'+
                    ' PSA_DATESORTIE="'+UsdateTime(Idate1900)+'" OR'+
                    ' PSA_DATESORTIE IS NULL) AND'+
                    ' PSA_DATEENTREE <="'+UsDateTime(DateArret)+'"';
      End
   else
      StDateArret:= '';
   End
Else
   StDateArret:='';

  //DEB PT14
  if Assigned(MonHabilitation) then
    CWhere := CWhere + ' AND ' + MonHabilitation.LeSQL;
  //FIN PT14

  SQL:=requete+CWhere+StDateArret+OrderBy;
  TFQRS1(Ecran).WhereSQL:=SQL;
end;
//FIN PT12


{-------------------------------------------------------------------------------
                             EDITION DES ACOMPTES
-------------------------------------------------------------------------------}
{ TOF_PGACOMPTE_ETAT }


procedure TOF_PGACOMPTE_ETAT.OnArgument(Arguments: string);
var
  Min, Max, DebPer, FinPer, ExerPerEncours: string;
  ok: Boolean;
  Edit: THEdit;
  Combo: ThValComboBox;
  Zone: TControl;
  Check: TCheckBox;
begin
  inherited;
  //Valeur par défaut
  Edit := ThEdit(getcontrol('DOSSIER'));
  if Edit <> nil then Edit.text := GetParamSoc('SO_LIBELLE'); //PT5

  Check := TCheckBox(GetControl('CKEURO'));
  if Check <> nil then Check.Checked := VH_Paie.PGTenueEuro;

  RecupMinMaxTablette('PG', 'SALARIES', 'PSA_SALARIE', Min, Max);
  Edit := ThEdit(getcontrol('PSD_SALARIE'));
  if Edit <> nil then
  begin
    Edit.text := Min;
    Edit.OnExit := ExitEdit;
  end;
  Edit := ThEdit(getcontrol('PSD_SALARIE_'));
  if Edit <> nil then
  begin
    Edit.text := Max;
    Edit.OnExit := ExitEdit;
  end;
  { DEB PT9 }
  RecupMinMaxTablette('PG', 'ETABLISS', 'ET_ETABLISSEMENT', Min, Max); //PT11
  Edit := ThEdit(getcontrol('PSD_ETABLISSEMENT'));
  if Edit <> nil then Edit.text := Min;
  Edit := ThEdit(getcontrol('PSD_ETABLISSEMENT_'));
  if Edit <> nil then Edit.text := Max;
  { FIN PT9 }
  SetControlText('PSD_DATEPAIEMENT', '01/01/1900');
  SetControlText('PSD_DATEPAIEMENT_', '01/01/2010');
  ok := RendPeriodeEnCours(ExerPerEncours, DebPer, FinPer);
  if Ok = True then
  begin
    SetControlText('PSD_DATEDEBUT', DebPer);
    SetControlText('PSD_DATEDEBUT_', FinPer);
  end;
  Combo := ThValComboBox(getcontrol('PSA_PAIACOMPTE'));
  if Combo <> nil then Combo.OnChange := ChangeModeReglement;
  Zone := ThValComboBox(getcontrol('PSA_PAIACOMPTE'));
  InitialiseCombo(Zone);
  { DEB PT9 Gestionnaire d'évènements }
  Check := TCheckBox(GetControl('CALPHA'));
  if Check <> nil then Check.OnClick := OnClickCritere;
  Check := TCheckBox(GetControl('CETAB'));
  if Check <> nil then Check.OnClick := OnClickCritere;
  Check := TCheckBox(GetControl('CMONT'));
  if Check <> nil then Check.OnClick := OnClickCritere;
  OnClickCritere(nil);
  { FIN PT9 }

end;

procedure TOF_PGACOMPTE_ETAT.ChangeModeReglement(Sender: TObject);
var
  ok: Boolean;
  DebPer, FinPer, ExerPerEncours: string;
begin
  if GetControlText('PSA_PAIACOMPTE') = '008' then
  begin
    SetControlEnabled('PSD_DATEPAIEMENT', True);
    SetControlEnabled('PSD_DATEPAIEMENT_', True);
    ok := RendPeriodeEnCours(ExerPerEncours, DebPer, FinPer);
    if Ok = True then
    begin
      SetControlText('PSD_DATEPAIEMENT', DebPer);
      SetControlText('PSD_DATEPAIEMENT_', FinPer);
    end;
  end
  else
  begin
    SetControlEnabled('PSD_DATEPAIEMENT', False);
    SetControlEnabled('PSD_DATEPAIEMENT_', False);
    SetControlText('PSD_DATEPAIEMENT', '01/01/1900');
    SetControlText('PSD_DATEPAIEMENT_', '01/01/2010');
  end;
end;
procedure TOF_PGACOMPTE_ETAT.ExitEdit(Sender: TObject);
var
  edit: thedit;
begin
  edit := THEdit(Sender);
  if edit <> nil then
    if edit.text <> '' then
      if (VH_Paie.PgTypeNumSal = 'NUM') and (length(Edit.text) < 11) and (isnumeric(edit.text)) then
        edit.text := AffectDefautCode(edit, 10);
end;

{ DEB PT9 }
procedure TOF_PGACOMPTE_ETAT.OnClickCritere(Sender: Tobject);
begin
  SetControlText('XX_RUPTURE1', '');
  if (PGisOracle) then   { PT13 }
    SetControlText('XX_ORDERBY', ' To_Char (PSD_DATEPAIEMENT,"YY"),To_Char (PSD_DATEPAIEMENT,"MM"),PSA_PAIACOMPTE,PSD_SALARIE') {PT10 PSD_DATEDEBUT,R_DOMICILIATION,}
  else
    if (PGisMssql) or (PGisSYBASE) then  { PT13 }
    SetControlText('XX_ORDERBY', ' DATEPART(YEAR,PSD_DATEPAIEMENT),DATEPART(MONTH,PSD_DATEPAIEMENT),PSA_PAIACOMPTE,PSD_SALARIE ') {PT10 PSD_DATEDEBUT,R_DOMICILIATION,}
  else
    SetControlText('XX_ORDERBY', ' YEAR(PSD_DATEPAIEMENT),MONTH(PSD_DATEPAIEMENT),PSA_PAIACOMPTE,PSD_SALARIE '); {PT10 PSD_DATEDEBUT,R_DOMICILIATION,}
  if GetControlText('CETAB') = 'X' then SetControlText('XX_RUPTURE1', 'PSD_ETABLISSEMENT');
  if GetControlText('CALPHA') = 'X' then
  begin
    SetControlChecked('CMONT', False);
    SetControlEnabled('CMONT', False);
    if (PGisOracle) then    { PT13 }
      SetControlText('XX_ORDERBY', ' To_Char (PSD_DATEPAIEMENT,"YY"),To_Char (PSD_DATEPAIEMENT,"MM"),PSA_PAIACOMPTE,PSA_LIBELLE,PSD_SALARIE') {PT10 PSD_DATEDEBUT,R_DOMICILIATION,}
    else
      if (PGisMssql) or (PGisSYBASE) then  { PT13 }
      SetControlText('XX_ORDERBY', ' DATEPART(YEAR,PSD_DATEPAIEMENT),DATEPART(MONTH,PSD_DATEPAIEMENT),PSA_PAIACOMPTE,PSA_LIBELLE,PSD_SALARIE ')
        {PT10 PSD_DATEDEBUT,R_DOMICILIATION,}
    else
      SetControlText('XX_ORDERBY', ' YEAR(PSD_DATEPAIEMENT),MONTH(PSD_DATEPAIEMENT),PSA_PAIACOMPTE,PSA_LIBELLE,PSD_SALARIE '); {PT10 PSD_DATEDEBUT,R_DOMICILIATION,}
  end else SetControlEnabled('CMONT', True);
  if GetControlText('CMONT') = 'X' then
  begin
    SetControlChecked('CALPHA', False);
    SetControlEnabled('CALPHA', False);
    if (PGisOracle) then  { PT13 }
      SetControlText('XX_ORDERBY', ' To_Char (PSD_DATEPAIEMENT,"YY"),To_Char (PSD_DATEPAIEMENT,"MM"),PSA_PAIACOMPTE,PSD_MONTANT DESC,PSD_SALARIE')
        {PT10 PSD_DATEDEBUT,R_DOMICILIATION,}
    else
      if (PGisMssql) or (PGisSYBASE) then  { PT13 }
      SetControlText('XX_ORDERBY', ' DATEPART(YEAR,PSD_DATEPAIEMENT),DATEPART(MONTH,PSD_DATEPAIEMENT),PSA_PAIACOMPTE,PSD_MONTANT DESC,PSD_SALARIE ')
        {PT10 PSD_DATEDEBUT,R_DOMICILIATION,}
    else
      SetControlText('XX_ORDERBY', ' YEAR(PSD_DATEPAIEMENT),MONTH(PSD_DATEPAIEMENT),PSA_PAIACOMPTE,PSD_MONTANT DESC,PSD_SALARIE '); {PT10 PSD_DATEDEBUT,R_DOMICILIATION,}
  end else SetControlEnabled('CALPHA', True);

  if Sender <> nil then
    if TCheckBox(Sender).Name = 'CALPHA' then
      AffectCritereAlpha(Ecran, TCheckBox(Sender).Checked, 'PSD_SALARIE', 'PSA_LIBELLE');
end;
{ FIN PT9 }

initialization
  registerclasses([TOF_PGMODEREGLE_ETAT, TOF_PGRIB_ETAT, TOF_PGACOMPTE_ETAT]);
end.

