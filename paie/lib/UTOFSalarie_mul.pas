{***********UNITE*************************************************
Auteur  ...... : PH
Créé le ...... : 11/09/2001
Modifié le ... :   /  /
Description .. : Unit de gestion du multi critère salariés
Mots clefs ... : PAIE
*****************************************************************
PT1 10/01/2002 SB V571 Restriction de la liste pour les salariés CP
PT2 29/01/2002 SB V571 Fiche de bug n°308 : Liste salarié date arrêté
PT3 31/01/2002 SB V571 Fiche de bug n°82 : Ajout CodeStat critère de selection
PT4 30/04/2002 JL V571 Ajout gestion bouton pour récupération intérimaire
PT5 13/05/2002 PH V575 Ajout controle date entrée par rapport à la date d'arreté
                  FQ 10124
PT6 28/05/2002 SB V582 Annulation PT2
PT7 12/12/2002 SB V591 FQ 10362 Affichage des salariés sortis sur coche
PT8 07/01/2003 SB V591 FQ 10430 Correction PT7
PT9 15/09/2003 SB V_421 Prise en compte autre cas
PT10 30/10/2003 SB V_50 FQ 10826 Design
PT11 12/05/2004 PH V_50 FQ 10973 REstriction accès salarié
PT12 16/09/2005 SB V_65 FQ 12516 Ajout Tof Mul Banque Salarié + exit salarié
PT13 02/10/2006 PH V_70 Mise en place des controles Paie 50
PT14 28/05/2007 JL V_720 Ajout TOF pour consultation bulletin dans fiche salarié
PT15 06/08/2007 JL V_80 FQ 14167 Gestion concept modif salarié
PT16 04/12/2007 FL V_81 Adaptation pour la fiche SALARIE_MULDOS : salariés multidossier
PT17 14/12/2007 FL V_81 Externalisation du lancement de l'application
}

unit UTOFSalarie_mul;

interface
uses  StdCtrls,Controls,Classes,sysutils,
      {$IFNDEF EAGLCLIENT}
      db,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}Mul,  HDB,Fe_Main,
      {$ELSE}
      eMul,UTOB,MainEAGL,
      {$ENDIF}
      HCtrls, HEnt1, HTB97, HMsgBox, UTOF, HQRY, ParamSoc, PGOutils, Forms, Windows;

Type
    TOF_Salarie_mul = Class (TOF)
       procedure OnArgument (stArgument: String);       Override;
       procedure ExitEdit(Sender: TObject);
       procedure OnLoad ;                               Override;
       procedure OnClose;                               Override;       //PT16
    private
       Argument,StTemp : string;
       ProcessInfo : TProcessInformation;       // Pour l'application lancée une seconde fois //PT16
       procedure OnClickSalarieSortie(Sender: TObject);
       {$IFDEF CPS1}
       procedure BtInsertclick(sender : tobject);
       {$ENDIF}
       //procedure LanceApplication (Dossier, Salarie : String); //PT16 //PT17
       procedure OuvreSalarieDos  (Sender: TObject);           //PT16
    END ;

    { DEB PT12 }
    TOF_MUL_SALARIEBQ  = Class (TOF)
       procedure OnArgument (stArgument: String);        Override;
       procedure OnExitSalarie(Sender: TObject);
    END;
    { FIN PT12 }

    { DEB PT14 }
    TOF_PGBULCONSULT = Class (TOF)
       procedure OnArgument (S : String ) ; override ;
       procedure OnLoad ;                   Override;
    private
       procedure BImprimerClick(Sender: TObject);
    END;
    { FIN PT14 }

implementation
uses P5Def, EntPaie, PgOutils2,shellapi,ed_tools;

procedure TOF_Salarie_mul.OnArgument (stArgument: String);
var
  Defaut: THEdit;
  Check: TCheckBox;
  Num: Integer;
  Q: Tquery;
  {$IFDEF CPS1}
  Bt: ttoolbarbutton97; // PT13
  {$ENDIF}
begin
    Argument :=stArgument; StTemp:='';
    // PT11 12/05/2004 PH V_50 FQ 10973 REstriction accès salarié
    if (Copy (Argument, 1, 6) = 'ACTION')  then
    begin
        if (Copy (Argument, 8, 4) = 'CONS') then SetControlText ('ACTION', 'CONSULTATION') ;
    end;
    // FIN PT11
    //DEB PT8
    SetControlvisible('DATEARRET',True);
    SetControlvisible('TDATEARRET',True);
    SetControlEnabled('DATEARRET',False);
    SetControlEnabled('TDATEARRET',False);
    //FIN PT8
    //DEB PT7
    Check:=TCheckBox(GetControl('CKSORTIE'));
    if Check=nil then
    Begin
        //Deb PT6
        SetControlVisible('DATEARRET',False);
        SetControlVisible('TDATEARRET',False);
        //FIN PT6
    End
    else
    Check.OnClick:=OnClickSalarieSortie;
    //FIN PT7
    For Num := 1 to VH_Paie.PGNbreStatOrg do
    begin
        if Num >4 then Break;
        VisibiliteChampSalarie (IntToStr(Num),GetControl ('PSA_TRAVAILN'+IntToStr(Num)),GetControl ('TPSA_TRAVAILN'+IntToStr(Num)));
    end;
    VisibiliteStat (GetControl ('PSA_CODESTAT'),GetControl ('TPSA_CODESTAT')) ; //PT3
    Defaut:=ThEdit(getcontrol('PSA_SALARIE'));
    If Defaut<>nil then Defaut.OnExit:=ExitEdit;
    if argument='CP' then //DEB PT1
    Begin
        //   SetControlText('XX_WHERE','');
        Q:=OpenSql('SELECT ETB_ETABLISSEMENT '+
        'FROM ETABCOMPL WHERE ETB_CONGESPAYES="X" ',True);
        StTemp:='(';
        While not Q.eof do
        Begin
            StTemp:=StTemp+'PSA_ETABLISSEMENT="'+Q.findField('ETB_ETABLISSEMENT').AsString+'" OR ';
            Q.next;
        End;
        if Length(StTemp)<10 then StTemp:='' else StTemp:=Copy(StTemp,1,Length(StTemp)-3)+')';
        SetControlText('XX_WHERE',StTemp);
        Ferme(Q);
    End;                //FIN PT1
    if (Argument = 'S') And (Ecran.Name <> 'SALARIE_MULDOS') then //DEBUT PT4 //PT16
    begin
        if (VH_Paie.PGInterimaires) and (VH_Paie.PGCodeInterim) then SetControlVisible('BRECUPINT',True);
    end;                       //FIN PT4
    {$IFDEF CPS1} // PT13
    if (V_PGI.ModePCL <> '1') AND (NOT PGControlBL (Paie1550)) then
    begin
        bt := ttoolbarbutton97(getcontrol('BInsert'));
        if bt <> nil then bt.onclick := btInsertclick;
    end;
    {$ENDIF}
    //DEBUT PT15
    if (Ecran.Name = 'SALARIE_MULDOS') Or (not JaiLeDroitTag(200001)) then //PT16
    begin
        SetControlText('ACTION', 'CONSULTATION');
        SetControlVisible('BINSERT',False);
    end;
    //FIN PT15
    //PT16 - Début
    if Ecran.Name = 'SALARIE_MULDOS' then
    begin
        {$IFNDEF EAGLCLIENT}THDBGrid{$ELSE}THGrid{$ENDIF}(GetControl('Fliste')).OnDblClick := OuvreSalarieDos;
    end;
    //PT16 - Fin
end;

procedure TOF_Salarie_mul.ExitEdit(Sender: TObject);
var edit : thedit;
begin
  edit := THEdit(Sender);
  if edit <> nil then //AffectDefautCode que si gestion du code salarié en Numérique
    if (VH_Paie.PgTypeNumSal = 'NUM') and (length(Edit.text) < 11) and (isnumeric(edit.text)) then
      edit.text := AffectDefautCode(edit, 10);
end;



procedure TOF_Salarie_mul.OnLoad ;
var
  Action: THEdit;
  sal: thedit;
  DateArret: TdateTime;
  StDateArret: string;
begin
  Action := thedit(getcontrol('ACTION'));
  if Action <> nil then
    Setcontrolvisible('BINSERT', Action.text <> 'CONSULTATION');

  if Argument = 'C' then
    Ecran.Caption := 'Saisie des congés payés';
  if Argument = 'SA' then
    Ecran.Caption := 'Saisie des absences';
  if Argument = 'CA' then
    Ecran.Caption := 'Consultation des absences par salarié';
  if Argument = 'GC' then
    Ecran.Caption := 'Gestion des contrats par salarié';
// PT9 15/09/2003 SB V_421 Prise en compte autre cas
  if Argument = 'CT' then
    Ecran.Caption := 'Gestion des compétences par salarié';
  { DEB PT10 }
  if Argument = 'CP' then
  Begin
    Ecran.Caption := 'Paramétrage CP salarié';
    Setcontrolvisible('BINSERT', False);
  End;
  { FIN PT10 }
  UpdateCaption(TFmul(Ecran));
  if (Argument = 'SA') or (Argument = 'CA') then
  begin
    sal := thedit(getcontrol('PSA_SALARIE'));
    if sal <> nil then
    begin
      sal.ElipsisButton := true;
      sal.datatype := 'PGSALARIE';
    end;
  end;
  SetControlText('XX_WHERE', StTemp);
  if TCheckBox(GetControl('CKSORTIE')) <> nil then
  Begin
  {PT6 mise en commentaire}
    if (GetControlText('CKSORTIE') = 'X') and (IsValidDate(GetControlText('DATEARRET'))) then //DEB PT2
    Begin
      DateArret := StrtoDate(GetControlText('DATEARRET'));
      StDateArret := ' AND (PSA_DATESORTIE>="' + UsDateTime(DateArret) + '" OR PSA_DATESORTIE="' + UsdateTime(Idate1900) + '" OR PSA_DATESORTIE IS NULL) ';
     // PT5 13/05/2002 PH V575 Ajout controle date entrée par rapport à la date d'arreté
      StDateArret := StDateArret + ' AND PSA_DATEENTREE <="' + UsDateTime(DateArret) + '"';
      SetControlText('XX_WHERE', StTemp + StDateArret);
    End //FIN PT2
    Else
      SetControlText('XX_WHERE', StTemp);
  End
  Else
    StDateArret := '';
end;

procedure TOF_Salarie_mul.OnClickSalarieSortie(Sender: TObject);
begin
//DEB PT7
  SetControlenabled('DATEARRET', (GetControltext('CKSORTIE') = 'X'));
  SetControlenabled('TDATEARRET', (GetControltext('CKSORTIE') = 'X'));
//FIN PT7
end;

{ TOF_MUL_SALARIEBQ }

{ DEB PT12 }
procedure TOF_MUL_SALARIEBQ.OnArgument(stArgument: String);
Var Defaut : ThEdit;
begin
  inherited;
  Defaut := ThEdit(getcontrol('PSA_SALARIE'));
  if Defaut <> nil then Defaut.OnExit := OnExitSalarie;
end;

procedure TOF_MUL_SALARIEBQ.OnExitSalarie(Sender: TObject);
var edit : thedit;
begin
  edit := THEdit(Sender);
  if edit <> nil then //AffectDefautCode que si gestion du code salarié en Numérique
    if (VH_Paie.PgTypeNumSal = 'NUM') and (length(Edit.text) < 11) and (isnumeric(edit.text)) then
      edit.text := AffectDefautCode(edit, 10);
end;
{ FIN PT12 }

// DEB PT13

{$IFDEF CPS1}
procedure TOF_Salarie_mul.BtInsertclick(sender: tobject);
begin
  PGIBOX ('Votre licence ne vous permet plus de créer de salariés.', Ecran.Caption);
end;
{$ENDIF}
// FIN PT13


{ DEB PT14 }
procedure TOF_PGBULCONSULT.OnArgument (S : String ) ;
var Salarie : String;
  vcbxMois, vcbxAnnee: THValComboBox;
  MoisE, AnneeE, ComboExer: string;
  DebExer, FinExer: TDateTime;
{$IFNDEF EAGLCLIENT}
  Grille: THDBGrid;
{$ELSE}
  Grille: THGrid;
{$ENDIF}
begin
  Inherited ;
  Salarie := ReadTokenPipe(S,';');
  SetControlText('PPU_SALARIE',Salarie);
  vcbxMois := THValComboBox(GetControl('CBXMOIS'));
  vcbxAnnee := THValComboBox(GetControl('CBXANNEE'));
  if RendExerSocialEnCours(MoisE, AnneeE, ComboExer, DebExer, FinExer) = TRUE then
  begin
    if vcbxMois <> nil then vcbxMois.value := MoisE;
    if vcbxAnnee <> nil then vcbxAnnee.value := ComboExer;
  end;
   {$IFNDEF EAGLCLIENT}
   Grille := THDBGrid (GetControl ('Fliste'));
   {$ELSE}
   Grille := THGrid (GetControl ('Fliste'));
   {$ENDIF}
   grille.OnDblClick := BImprimerClick;
end;

procedure TOF_PGBULCONSULT.OnLoad ;
var Mois,Annee : String;
  WMois, WAnnee: WORD;
  LaDate, DebutBulletin, FinBulletin: TDateTime;
begin
  inherited;
  Mois := GetControlText('CBXMOIS');
  Annee := GetControlText('CBXANNEE');
  If Annee <> '' then Annee := RechDom('PGANNEESOCIALE', Annee, FALSE);
  If Annee = '' then
  begin
    If Mois <> '' then SetControlText('CBXMOIS','');
    SetControlText('XX_WHERE','');
  end
  else
  begin
    If Mois <> '' then WMois := Trunc(StrToInt(Mois))
    else WMois := 1;
    WAnnee := Trunc(StrToInt(Annee));
    LaDate := EncodeDate(Wannee, Wmois, 1);
    DebutBulletin := LaDate;
    if Mois <> '' then LaDate := FINDEMOIS(LaDate)
    else laDate := EncodeDate(Wannee, 12, 31);
    FinBulletin := LaDate;
    SetControlText('XX_WHERE','PPU_DATEFIN >="' + UsDateTime(DebutBulletin) + '" AND PPU_DATEFIN <="' + UsDateTime(FinBulletin) + '"');
  end;

end;


procedure TOF_PGBULCONSULT.BImprimerClick(Sender: TObject);
var
  St: string;
  Q_Mul : THQuery;
  etab,CodeSalarie : String;
  DateD,DateF : TDateTime;
begin
  Q_Mul := THQuery(Ecran.FindComponent('Q'));
  CodeSalarie := Q_MUL.FindField('PPU_SALARIE').AsString;
  Etab := Q_MUL.FindField('PPU_ETABLISSEMENT').AsString;
  DateD := Q_MUL.FindField('PPU_DATEDEBUT').AsDateTime;
  DateF := Q_MUL.FindField('PPU_DATEFIN').AsDateTime;
  St :=  Etab + ';' + CodeSalarie + ';' + DateToStr(DateD) + ';' + DateToStr(DateF);
{$IFNDEF EAGLSERVER}
  if (Etab <> '') and (CodeSalarie <> '') and (DateD > 0) and (DateF > 0) then
    // Si on ne gère pas le mode d'édition
    if not GetParamSocSecur('SO_PGGESTORIDUPSPE', False) then
      AglLanceFiche('PAY', 'EDITBUL_ETAT', '', '', St + ';BULL;' + '-')
    else
    begin
      // Vérifier que l'utilisateur a au moins le droit d'éditer un duplicata ou un specimen
      // A ce niveau, on ne propose pas l'édition d'un original
      if JaiLeDroitTag(42314) or JaiLeDroitTag(42316) then
      begin
        AglLanceFiche('PAY', 'MODEEDT_BUL', '', '', 'SAISBUL');
        if PGModeEdition <> '' then
          AglLanceFiche('PAY', 'EDITBUL_ETAT', '', '', St + ';BULL;' + '-' + ';' + PGModeEdition);
      end
      else
        PGIBox(TraduireMemoire('Vous n''êtes pas autorisé à lancer une édition de bulletins'));
    end;
{$ENDIF}
end;

//PT16 - Début
procedure TOF_SALARIE_MUL.OuvreSalarieDos(Sender: TObject);
var LaBase, LaBaseOrig, CodeSalarie: string;
  Nom     : String;
begin
    LaBaseOrig := V_PGI.CurrentAlias;

    {$IFDEF EAGLCLIENT}
    TFMul(Ecran).Q.TQ.Seek(TFMul(Ecran).FListe.Row - 1);
    {$ENDIF}

    If GetControlText('MULTIDOSSIER') <> '' Then
    Begin
        Try
          LaBase    := TFmul(Ecran).Q.FindField('SYSDOSSIER').AsString;
        Except
          LaBase    := LaBaseOrig;
        End;
    End
    Else
        LaBase      := LaBaseOrig;

    CodeSalarie := TFmul(Ecran).Q.FindField('PSA_SALARIE').AsString;
    Nom         := TFmul(Ecran).Q.FindField('PSA_LIBELLE').AsString + ' ' + TFmul(Ecran).Q.FindField('PSA_PRENOM').AsString;

    If (LaBase <> '') And (LaBase <> LaBaseOrig) Then
    Begin    	
        // On bloque tout pour ne pas que l'utilisateur puisse intéragir
        SetControlEnabled('BOuvrir',   False);  //PT17
        SetControlEnabled('BAnnuler',  False);
        SetControlEnabled('BSelectAll',False);
        
        // Fenêtre d'attente et blocage de l'application tant que le traitement ne sera pas terminé
        InitMoveProgressForm(Application.MainForm, TraduireMemoire('Veuillez patienter svp...'), Format(TraduireMemoire('Chargement du salarié %s, Base %s'),[Nom, LaBase]), 4,  False, True);

        // Lancement de l'exécutable
        LanceApplication (ProcessInfo, LaBase, '/FICHE=SALARIE&' + CodeSalarie + '&ACTION=CONSULTATION;MONOFICHE', True); //PT17

        FiniMoveProgressForm;
        
        // On débloque les boutons
        SetControlEnabled('BOuvrir',   True);  //PT17
        SetControlEnabled('BAnnuler',  True);
        SetControlEnabled('BSelectAll',True);
    End
    Else
    Begin
        AGLLanceFiche('PAY', 'SALARIE', '', CodeSalarie, 'ACTION=MODIFICATION;MONOFICHE');
    End;
end;

procedure  TOF_SALARIE_MUL.OnClose;
begin
    // Si on a ouvert la fiche salarié d'un autre dossier, on bloque la fermeture de la fenêtre mul
    if  (ProcessInfo.hProcess <> 0) Then
        LastError := 1
    Else
        LastError := 0;
end;
//PT16 - Fin

Initialization
registerclasses([TOF_Salarie_mul,TOF_MUL_SALARIEBQ,TOF_PGBULCONSULT]) ;

end.



