{***********UNITE*************************************************
Auteur  ...... : René Marsella
Créé le ...... : 20/02/2007
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : PGDUESAISIE ()
Mots clefs ... : TOF;PGDUESAISIE
*****************************************************************}
Unit UTOFPGDUESAISIE ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,
{$else}
     eMul,
{$ENDIF}
     forms, sysutils, ComCtrls, HCtrls, HEnt1, HMsgBox,
     vierge, UTOF ,UTOB ,ParamSoc , HQry, entpaie, PGoutils2,
     UtobDebug ;

Type
  TOF_PGDUESAISIE = Class (TOF)
    private
      Demande ,Sal_Matr :String;
      TobDUE  : Tob;
      CtrlDUE, FaireMaj : Boolean;
      Nordre  : Integer;

      Procedure DUE_REMPLIR_PAR_DEFAUT(Sender: TObject);
      Procedure DUE_RENDRE_CHAMP_SAISISSABLE(OuiNon : Boolean);
      Procedure DUE_RENSEIGNER_CHAMP;
      Procedure DUE_ENREGISTRER_EN_TABLE;
      Procedure DUE_LIRE_CHAMP_TEDIT(Sender: TObject);
      Procedure DUE_LIRE_CHAMP_TCOMBO(Sender: TObject);
      Procedure DUE_LIRE_CHAMP_TBOX(Sender: TObject);
    public
      procedure OnNew                    ; override ;
      procedure OnDelete                 ; override ;
      procedure OnUpdate                 ; override ;
      procedure OnLoad                   ; override ;
      procedure OnArgument (S : String ) ; override ;
      procedure OnDisplay                ; override ;
      procedure OnClose                  ; override ;
      procedure OnCancel                 ; override ;
  end ;

Implementation

procedure TOF_PGDUESAISIE.OnNew ;
begin
  Inherited ;
  PGIBOX('OnNew,'+TFVierge(Ecran).Caption+'');
end ;

procedure TOF_PGDUESAISIE.OnDelete ;
begin
  Inherited ;
  PGIBOX('OnDelete,'+TFVierge(Ecran).Caption+'');
end ;

procedure TOF_PGDUESAISIE.OnUpdate;
begin
  Inherited ;
  If FaireMaj = True Then DUE_ENREGISTRER_EN_TABLE;
end ;

procedure TOF_PGDUESAISIE.OnLoad ;
begin
  Inherited ;
  //PGIBOX('OnLoad,'+TFVierge(Ecran).Caption+'');
end ;

Procedure TOF_PGDUESAISIE.OnArgument (S : String ) ;
var
  PDUE: TPageControl;
  Defaut, TEdit : THEdit;
  TCombo : THValComboBox;
  Tbox   : THCheckBox;
  SQL, CTRL, Nom  : String;
  Q :TQuery;
  I : Integer;

begin
  Inherited ;
  Demande := Trim(S);
  PDUE := TPageControl(GetControl('PAGEDUE'));
  If PDUE = nil Then Exit;

  Sal_Matr := '';
  CtrlDUE  := False;
  FaireMaj := False;

  SQL := 'SELECT * FROM DUELEXIQUE';
  Q := OpenSQL(SQL,True);
  TobDUE := Tob.Create('DUELEXIQUE',Nil,-1);
  TobDUE.LoadDetailDB('DUELEXIQUE','','',Q,False);
  Ferme(Q);

  For i := 0 To TobDUE.Detail.Count -1 do
  Begin
    TobDUE.Detail[i].AddChampSupValeur('DONNEE','');
    TobDUE.Detail[i].AddChampSupValeur('DONNEEAFFICH','');
  End;

  If Demande = 'CREATION' Then
  Begin
     TFVierge(Ecran).Caption := 'Saisie d''une nouvelle DUE';
     UpdateCaption(TFVierge(Ecran));
     PDUE.ActivePageIndex := 1;
     SetControlEnabled('PUD_SALARIE',TRUE);
     SetFocusControl('PUD_SALARIE');
     Defaut := ThEdit(Getcontrol('PUD_SALARIE'));
     If Defaut <> nil Then Defaut.OnExit := DUE_REMPLIR_PAR_DEFAUT;
  End;

  For i := 0 To TobDUE.Detail.Count -1 do
  Begin
    Nom  := TobDUE.Detail[i].GetValue('PUL_NOMZONE');
    CTRL := TobDUE.Detail[i].GetValue('PUL_DUEVALEUR');
    If (Pos('TEdit',CTRL) <> 0) Then
    Begin
       TEdit := ThEdit(Getcontrol(Nom));
       If TEdit <> nil Then TEdit.OnExit := DUE_LIRE_CHAMP_TEDIT;
    End;
    If (Pos('TCombo',CTRL) <> 0) Then
    Begin
       TCombo := THValComboBox(Getcontrol(Nom));
       If TCombo <> nil Then TCombo.OnExit := DUE_LIRE_CHAMP_TCOMBO;
    End;
    If (Pos('Tbox',CTRL) <> 0) Then
    Begin
       Tbox := THCheckBox(Getcontrol(Nom));
       If Tbox <> nil Then Tbox.OnClick := DUE_LIRE_CHAMP_TBOX;
    End;
  End;
  Tbox := THCheckBox(Getcontrol('T_MT_DCL'));
  If Tbox <> nil Then Tbox.OnClick := DUE_LIRE_CHAMP_TBOX;



  //SetControlEnabled('EDUEEMETTEUR',OuiNon);
  //SetControlEnabled('EFAXCORRESP',OuiNon);
End;

//===========//
procedure TOF_PGDUESAISIE.OnClose ;
begin
  Inherited ;
  //PGIBOX('OnClose,'+TFVierge(Ecran).Caption+'');
  FreeAndNil(TobDUE);
End ;

procedure TOF_PGDUESAISIE.OnDisplay () ;
begin
  Inherited ;
  PGIBOX('OnDisplay,'+TFVierge(Ecran).Caption+'');
end ;

procedure TOF_PGDUESAISIE.OnCancel () ;
begin
  Inherited ;
  PGIBOX('OnCancel,'+TFVierge(Ecran).Caption+'');
end ;

//===========//
Procedure TOF_PGDUESAISIE.DUE_REMPLIR_PAR_DEFAUT(Sender: TObject) ;
var
  SQL, Nchamp, Nemetteur, Identifiant, AdrAR : String;
  NomResp, TelResp, TChamp, Oblig : String;
  DatDeb, DatFin : TdateTime;
  Q :TQuery;
  I, NbJ, SocSal, EtbSal, Long, MTrunc : Integer;
  TobSAL, TobEtb, TobEtbC, TobCt : Tob;
  Edit : THEdit;

Begin
  NomResp := '';
  TelResp := '';
  Identifiant := '';
  AdrAR  := '';
  EtbSal := 0;
  SocSal := 0;
  Nordre := 0;

  Edit := THEdit(Sender);
  If Edit <> nil Then	//AffectDefautCode que si gestion du code salarié en Numérique
    If (VH_Paie.PgTypeNumSal='NUM') And (Length(Edit.text) < 11) And (Isnumeric(Edit.text)) Then
       Edit.text := AffectDefautCode(Edit,10);

  If Sal_Matr = Edit.text Then Exit; //On est repassé sur le champ sans le modifier
  Sal_Matr := Edit.text;
  If Trim(Sal_Matr) = '' Then Exit;

  If Presence('SALARIES','PSA_SALARIE',Sal_Matr) = False Then
  Begin
    PGIBOX('Le salarié n''a pas été créé',''+TFVierge(Ecran).Caption+'');
    SetFocusControl('PUD_SALARIE');
    Exit;
  End;

  Nemetteur := GetParamSoc('SO_PGDUEEMETTEUR');
  If Trim(Nemetteur) <> '' Then
  Begin
     SetControlText('EDUEEMETTEUR',Nemetteur);
     SQL := 'SELECT PET_CONTACTDUCS,PET_TELDUCS,PET_DUEIDENT,PET_DUEADRAR FROM EMETTEURSOCIAL Where PET_EMETTSOC = "'+Nemetteur+'"';
     Q := OpenSQL(SQL,True);
     If not Q.eof Then
     Begin
        NomResp := Q.Fields[0].asstring;
        TelResp := Q.Fields[1].asstring;
        Identifiant := Q.Fields[2].asstring;
        AdrAR := Q.Fields[3].asstring;
     End;
     Ferme(Q);
  End;

  SQL := 'SELECT * FROM SALARIES Where PSA_SALARIE = "'+Sal_Matr+'"';
  Q := OpenSQL(SQL,True);
  TobSAL := Tob.Create('le salarie',Nil,-1);
  TobSAL.LoadDetailDB('le salarie','','',Q,False);
  Ferme(Q);
  If TobSAL.Detail.Count = 0 Then
  Begin
    TobSAL.Free;
    Exit;
  End;

  SQL := 'SELECT * FROM ETABLISS Where ET_ETABLISSEMENT = "'+TobSAL.Detail[0].GetValue('PSA_ETABLISSEMENT')+'"';
  Q := OpenSQL(SQL,True);
  TobEtb := Tob.Create('Letab',Nil,-1);
  TobEtb.LoadDetailDB('Letab','','',Q,False);
  Ferme(Q);
  If TobEtb.Detail.Count = 0 Then
  Begin
    PGIBOX('L''établissement '+TobSAL.Detail[0].GetValue('PSA_ETABLISSEMENT')+' n''a pas été trouvé',''+TFVierge(Ecran).Caption+'');
    TobSAL.Free;
    TobEtb.Free;
    Exit;
  End;

  SQL := 'SELECT * FROM ETABCOMPL Where ETB_ETABLISSEMENT = "'+TobSAL.Detail[0].GetValue('PSA_ETABLISSEMENT')+'"';
  Q := OpenSQL(SQL,True);
  TobEtbC := Tob.Create('LetabComp',Nil,-1);
  TobEtbC.LoadDetailDB('LetabComp','','',Q,False);
  Ferme(Q);
  If TobEtbC.Detail.Count = 0 Then
  Begin
    PGIBOX('Le complément social de l''établissement '+TobSAL.Detail[0].GetValue('PSA_ETABLISSEMENT')+' n''a pas été trouvé',''+TFVierge(Ecran).Caption+'');
    TobSAL.Free;
    TobEtb.Free;
    TobEtbC.Free;
    Exit;
  End;

  SQL := 'SELECT * FROM CONTRATTRAVAIL Where PCI_SALARIE = "'+Sal_Matr+'" Order by PCI_ORDRE';
  Q := OpenSQL(SQL,True);
  TobCt := Tob.Create('Contrat',Nil,-1);
  TobCt.LoadDetailDB('Contrat','','',Q,False);
  Ferme(Q);
  If TobCt.Detail.Count = 0 Then
  Begin
    PGIBOX('Le contrat de travail pour le salarié matricule '+Sal_Matr+' n''a pas été créé',''+TFVierge(Ecran).Caption+'');
  End;

  SQL:='SELECT COUNT(PSA_SALARIE) FROM SALARIES';
  Q := OpenSQL(SQL,True);
  If not Q.eof Then
  Begin
    SocSal := Q.Fields[0].asInteger;
  End;
  Ferme(Q);
  SQL:='SELECT COUNT(PSA_SALARIE) FROM SALARIES Where PSA_ETABLISSEMENT = "'+TobSAL.Detail[0].GetValue('PSA_ETABLISSEMENT')+'"';
  Q := OpenSQL(SQL,True);
  If not Q.eof Then
  Begin
    EtbSal := Q.Fields[0].asInteger;
  End;
  Ferme(Q);

  SQL := 'SELECT MAX(PUD_ORDRE) FROM DUEDONNEE Where PUD_SALARIE = "'+Sal_Matr+'"';
  Q := OpenSQL(SQL,True);
  If not Q.eof Then
  Begin
    Nordre := Q.Fields[0].asInteger;
  End;
  Ferme(Q);
  Nordre := Nordre + 1;

  For i := 0 To TobDUE.Detail.Count -1 do
  Begin
     Nchamp := TobDUE.Detail[i].GetValue('PUL_NOMZONE');
     Long   := TobDUE.Detail[i].GetValue('PUL_LONGZONE');
     Oblig  := TobDUE.Detail[i].GetValue('PUL_OBLIGZONE');
     TChamp := '';

     If Nchamp = 'L_IDENT_EMT' Then TChamp := Identifiant;
     If Nchamp = 'N_FAX_AR_LOT' Then
     Begin
        TChamp := GetParamSoc('SO_PGDUEFAXRESP');
        If Trim(TChamp) = '' Then
        Begin
           ForceNumerique(TobEtb.Detail[0].GetValue('ET_FAX'),TChamp);
           TChamp := Copy(TChamp,1,Long);
        End;
     End;
     If Nchamp = 'C_ADR_RET_AR'   Then TChamp := AdrAR;
     If Nchamp = 'D_CREATION'     Then TChamp := DateToStr(Date);
     If Nchamp = 'H_CREATION'     Then TChamp := TimeToStr(Time);
     If Nchamp = 'N_SIRET'        Then TChamp := Copy(TobEtb.Detail[0].GetValue('ET_SIRET'),1,Long);
     If Nchamp = 'C_TYPE_DECL'    Then TChamp := GetParamSoc('SO_PGDUETYPEDECL');
     If Nchamp = 'L_RAISON_SOC_1' Then TChamp := Copy(TobEtb.Detail[0].GetValue('ET_LIBELLE'),1,Long);
     If Nchamp = 'C_NAF'          Then TChamp := Copy(TobEtb.Detail[0].GetValue('ET_APE'),1,Long);
     If Nchamp = 'L_ADR_EMP_1'    Then TChamp := Copy(TobEtb.Detail[0].GetValue('ET_ADRESSE1'),1,Long);
     If Nchamp = 'L_ADR_EMP_2'    Then TChamp := Copy(TobEtb.Detail[0].GetValue('ET_ADRESSE2'),1,Long);
     If Nchamp = 'C_POSTAL_EMP'   Then TChamp := Copy(TobEtb.Detail[0].GetValue('ET_CODEPOSTAL'),1,Long);
     If Nchamp = 'L_BUR_DIST_EMP' Then TChamp := Copy(TobEtb.Detail[0].GetValue('ET_VILLE'),1,Long);
     If Nchamp = 'L_NOM_PATRO_SAL' Then
     Begin
        TChamp := Copy(TobSAL.Detail[0].GetValue('PSA_NOMJF'),1,Long);
        If Trim(TChamp) = '' Then TChamp := Copy(TobSAL.Detail[0].GetValue('PSA_LIBELLE'),1,Long);
     End;
     If Nchamp = 'L_PRENOMS_SAL' Then TChamp := Copy(TobSAL.Detail[0].GetValue('PSA_PRENOM'),1,Long);
     If Nchamp = 'N_SECU_SOC'    Then TChamp := Copy(TobSAL.Detail[0].GetValue('PSA_NUMEROSS'),1,Long);
     If Nchamp = 'D_NAISSANCE_SAL'  Then TChamp := TobSAL.Detail[0].GetValue('PSA_DATENAISSANCE');
     If Nchamp = 'L_LIEU_NAISS_SAL' Then TChamp := Copy(TobSAL.Detail[0].GetValue('PSA_COMMUNENAISS'),1,Long);
     If Nchamp = 'L_PAYS_SAL'  Then TChamp := TobSAL.Detail[0].GetValue('PSA_PAYSNAISSANCE');
     If Nchamp = 'D_EMBAUCHE'  Then TChamp := TobSAL.Detail[0].GetValue('PSA_DATEENTREE');
     If Nchamp = 'H_EMBAUCHE'  Then TChamp := TobSAL.Detail[0].GetValue('PSA_HEUREMBAUCHE');
     If Nchamp = 'L_ACTIVITE'  Then TChamp := TobEtbC.Detail[0].GetValue('ETB_ACTIVITE');
     If Nchamp = 'C_FORME_JUR' Then TChamp := TobEtb.Detail[0].GetValue('ET_JURIDIQUE');
     If Nchamp = 'N_TEL_EMP'   Then
     Begin
        ForceNumerique(TobEtb.Detail[0].GetValue('ET_TELEPHONE'),TChamp);
        TChamp := Copy(TChamp,1,Long);
     End;
     If Nchamp = 'N_FAX_EMP' Then
     Begin
        ForceNumerique(TobEtb.Detail[0].GetValue('ET_FAX'),TChamp);
        TChamp := Copy(TChamp,1,Long);
     End;
     If Nchamp = 'L_ADR_CRESP_1'    Then TChamp := GetParamSoc('SO_PGDUEADR1RESP');
     If Nchamp = 'L_ADR_CRESP_2'    Then TChamp := GetParamSoc('SO_PGDUEADR2RESP');
     If Nchamp = 'C_POST_CRESP'     Then TChamp := GetParamSoc('SO_PGDUECPRESP');
     If Nchamp = 'L_BUR_DIST_CRESP' Then TChamp := GetParamSoc('SO_PGDUEVILLERESP');
     If Nchamp = 'C_CVLT_SAL'       Then TChamp := TobSAL.Detail[0].GetValue('PSA_CIVILITE');
     If Nchamp = 'L_NOM_EPX_SAL' Then
        If Trim(TobSAL.Detail[0].GetValue('PSA_NOMJF')) <> '' Then TChamp := Copy(TobSAL.Detail[0].GetValue('PSA_LIBELLE'),1,Long);
     If Nchamp = 'C_SEXE_SAL'  Then TChamp := TobSAL.Detail[0].GetValue('PSA_SEXE');
     If Nchamp = 'L_ADR_SAL_1' Then TChamp := Copy(TobSAL.Detail[0].GetValue('PSA_ADRESSE1'),1,Long);
     If Nchamp = 'L_ADR_SAL_2' Then TChamp := Copy(TobSAL.Detail[0].GetValue('PSA_ADRESSE2'),1,Long);
     If Nchamp = 'C_POST_SAL'  Then TChamp := Copy(TobSAL.Detail[0].GetValue('PSA_CODEPOSTAL'),1,Long);
     If Nchamp = 'L_BUR_DIST_SAL'   Then TChamp := Copy(TobSAL.Detail[0].GetValue('PSA_VILLE'),1,Long);
     If Nchamp = 'C_DEPT_NAISS_SAL' Then TChamp := Copy(TobSAL.Detail[0].GetValue('PSA_DEPTNAISSANCE'),1,Long);
     If Nchamp = 'L_NATION_SAL'     Then TChamp := TobSAL.Detail[0].GetValue('PSA_NATIONALITE');
     If (Nchamp = 'C_TYPE_CONTRAT') And (TobCt.Detail.Count > 0) Then
        TChamp := TobCt.Detail[0].GetValue('PCI_TYPECONTRAT');
     If (Nchamp = 'D_FIN_CDD') And (TobCt.Detail.Count > 0) Then
        TChamp := TobCt.Detail[0].GetValue('PCI_FINCONTRAT');
     If (Nchamp = 'N_HH_HEBDO') And (TobCt.Detail.Count > 0) Then
     Begin
        MTrunc := TobCt.Detail[0].GetValue('PCI_HORHEBDO');
        MTrunc := Trunc(MTrunc);
        TChamp := IntToStr(MTrunc);
     End;
     If (Nchamp = 'N_JJ_ESSAI') And (TobCt.Detail.Count > 0) Then
     Begin
        Nbj := 0;
        DatDeb := Idate1900;
        DatFin := Idate1900;
        If IsValidDate(TobCt.Detail[0].GetValue('PCI_ESSAIDEBUT')) Then DatDeb := TobCt.Detail[0].GetValue('PCI_ESSAIDEBUT');
        If IsValidDate(TobCt.Detail[0].GetValue('PCI_ESSAIFIN')) Then DatFin := TobCt.Detail[0].GetValue('PCI_ESSAIFIN');
        If (DatDeb > Idate1900) And (DatFin > Idate1900) And (DatFin >= DatDeb) Then
        Begin
           While DatDeb <= DatFin do
           Begin
              DatDeb := DatDeb + 1;
              NbJ := NbJ + 1;
           End;
        End;
        TChamp := IntToStr(NbJ);
     End;
     If (Nchamp = 'N_SALAIRE_MM') And (TobCt.Detail.Count > 0) Then
        TChamp := Copy(TobCt.Detail[0].GetValue('PCI_MONTANTCT'),1,Long);
     If Nchamp = 'L_NOM_RESP' Then
     Begin
        TChamp := GetParamSoc('SO_PGDUENOMRESP');
        If Trim(TChamp) = '' Then TChamp := NomResp;
     End;
     If Nchamp = 'N_TEL_RESP' Then
     Begin
        TChamp  := GetParamSoc('SO_PGDUETELRESP');
        If Trim(TChamp) = '' Then TChamp  := TelResp;
     End;
     If (Nchamp = 'N_HH_TRAVAIL') And (TobCt.Detail.Count > 0) Then
     Begin
        MTrunc := TobCt.Detail[0].GetValue('PCI_HORAIREMOIS');
        MTrunc := Trunc(MTrunc);
        TChamp := IntToStr(MTrunc);
     End;
     If Nchamp = 'N_EFFECTIF_ETAB' Then TChamp := IntToStr(EtbSal);
     If Nchamp = 'N_EFFECTIF_ENTR' Then TChamp := IntToStr(SocSal);
     If Nchamp = 'C_UNITE_HH' Then TChamp := 'M';
     If (Nchamp = 'I_PREMIERE_EMB') And (EtbSal = 0) Then TChamp := 'O';
     If (Nchamp = 'I_PMF5') And (EtbSal = 0) Then TChamp := 'O';

     If Trim(TChamp) <> '' Then
     Begin
        If Oblig = 'X' Then TChamp := UpperCase(TChamp);
        TobDUE.Detail[i].PutValue('DONNEEAFFICH',TChamp);
     End;
  End;
  TobSAL.Free;
  TobEtb.Free;
  TobEtbC.Free;
  TobCt.Free;
  //TOBDebug(TobDUE);
  DUE_RENDRE_CHAMP_SAISISSABLE(True);
  DUE_RENSEIGNER_CHAMP;
  FaireMaj := True;
End;

//===========//
Procedure TOF_PGDUESAISIE.DUE_LIRE_CHAMP_TEDIT(Sender: TObject) ;
var
  TEdit : THEdit;
  I : Integer;
  Nom : String;

Begin
  TEdit := THEdit(Sender);
  If TEdit <> nil Then
  Begin
     For i := 0 To TobDUE.Detail.Count -1 do
     Begin
        Nom  := TobDUE.Detail[i].GetValue('PUL_NOMZONE');
        If Tedit.name = Nom Then
        Begin
           If Tedit.name = 'C_MT_DCL' Then
           Begin
              If (Trim(Tedit.Text)='') And (GetControlText('T_MT_DCL')='X') Then Exit;
              If (Trim(Tedit.Text)<>'')  And (GetControlText('T_MT_DCL')='X') Then SetControlChecked('T_MT_DCL',False);
              TobDUE.Detail[i].PutValue('DONNEEAFFICH',Tedit.Text);
           End
           Else
           Begin
             If Tedit.name = 'C_URSSAF' Then TobDUE.Detail[2].PutValue('DONNEEAFFICH',Tedit.Text);
             TobDUE.Detail[i].PutValue('DONNEEAFFICH',Tedit.Text);
           End;
           Exit;
        End;
     End;
  End;
End;

//===========//
Procedure TOF_PGDUESAISIE.DUE_LIRE_CHAMP_TCOMBO(Sender: TObject) ;
var
  TCombo : THValComboBox;
  I : Integer;
  Nom  : String;

Begin
  TCombo := THValComboBox(Sender);
  If TCombo <> nil Then
  Begin
     For i := 0 To TobDUE.Detail.Count -1 do
     Begin
        Nom  := TobDUE.Detail[i].GetValue('PUL_NOMZONE');
        If TCombo.name = Nom Then
        Begin
           TobDUE.Detail[i].PutValue('DONNEEAFFICH',TCombo.Value);
           Exit;
        End;
     End;
  End;
End;

//===========//
Procedure TOF_PGDUESAISIE.DUE_LIRE_CHAMP_TBOX(Sender: TObject) ;
var
  Tbox : THCheckBox;
  I : Integer;
  Nom  : String;

Begin
  Tbox := THCheckBox(Sender);
  If Tbox <> nil Then
  Begin
     For i := 0 To TobDUE.Detail.Count -1 do
     Begin
        Nom  := TobDUE.Detail[i].GetValue('PUL_NOMZONE');
        If Tbox.name = Nom Then
        Begin
           If Tbox.Checked = False Then
              TobDUE.Detail[i].PutValue('DONNEEAFFICH',' ')
            Else
              TobDUE.Detail[i].PutValue('DONNEEAFFICH','O');
           Exit;
        End
        Else If (Tbox.name = 'T_MT_DCL') And (Nom = 'C_MT_DCL') Then
        Begin
           If Tbox.Checked = True Then
           Begin
              If (GetControlText('C_MT_DCL')<> '') Then SetControlText('C_MT_DCL','');
              TobDUE.Detail[i].PutValue('DONNEEAFFICH','999')
           End
           Else
              If (GetControlText('C_MT_DCL')= '') Then TobDUE.Detail[i].PutValue('DONNEEAFFICH','');
           Exit;
        End;
     End;
  End;
End;

//===========//
Procedure TOF_PGDUESAISIE.DUE_RENDRE_CHAMP_SAISISSABLE(OuiNon : Boolean);
Var
  I : Integer;
  CTRL, Nom, Oblig : String;
  TEdit  : THEdit;

Begin
  For i := 0 To TobDUE.Detail.Count -1 do
  Begin
    Nom  := TobDUE.Detail[i].GetValue('PUL_NOMZONE');
    CTRL := TobDUE.Detail[i].GetValue('PUL_DUEVALEUR');
    Oblig:= TobDUE.Detail[i].GetValue('PUL_OBLIGZONE');

    If (Pos('Snon',CTRL) <> 0) Then
    Begin
       SetControlEnabled(Nom,False);
       Continue;
    End;

    If (Pos('TEdit',CTRL) <> 0) Or (Pos('TCombo',CTRL) <> 0) Or (Pos('Tbox',CTRL) <> 0) Then
       SetControlEnabled(Nom,OuiNon);

    If (Pos('TEdit',CTRL) <> 0) And (Oblig = 'X') Then
    Begin
       TEdit := ThEdit(Getcontrol(Nom));
       If TEdit <> nil Then TEdit.CharCase := ecUpperCase;
    End;
  End;
  SetControlEnabled('T_MT_DCL',OuiNon);
  SetControlEnabled('EDUEEMETTEUR',OuiNon);
  SetControlEnabled('EFAXCORRESP',OuiNon);
End;

//===========//
Procedure TOF_PGDUESAISIE.DUE_RENSEIGNER_CHAMP;
Var
  I : Integer;
  SMessage, CTRL, Cbox : String;

Begin
  For i := 0 To TobDUE.Detail.Count -1 do
  Begin
    CTRL := TobDUE.Detail[i].GetValue('PUL_DUEVALEUR');
    If i=68 Then  //C_MT_DCL
    Begin
       Cbox := Trim(TobDUE.Detail[i].GetValue('DONNEEAFFICH'));
       If (Cbox <> '999') And (Cbox <> '') Then
       Begin
          SetControlChecked('T_MT_DCL',False);
          SetControlText(TobDUE.Detail[i].GetValue('PUL_NOMZONE'),TobDUE.Detail[i].GetValue('DONNEEAFFICH'));
       End
       Else If (Cbox = '999') Then
       Begin
          SetControlChecked('T_MT_DCL',True);
          SetControlText(TobDUE.Detail[i].GetValue('PUL_NOMZONE'),'');
       End
       Else
       Begin
          SetControlChecked('T_MT_DCL',False);
          SetControlText(TobDUE.Detail[i].GetValue('PUL_NOMZONE'),'');
       End;
       Continue;
    End;

    If (Pos('TEdit',CTRL) <> 0) Or (Pos('TCombo',CTRL) <> 0) Then
       SetControlText(TobDUE.Detail[i].GetValue('PUL_NOMZONE'),TobDUE.Detail[i].GetValue('DONNEEAFFICH'));
    If Pos('Tbox',CTRL) <> 0 Then
    Begin
       Cbox := Trim(TobDUE.Detail[i].GetValue('DONNEEAFFICH'));
       If Cbox = 'O' Then SetControlChecked(TobDUE.Detail[i].GetValue('PUL_NOMZONE'),True)
                     Else SetControlChecked(TobDUE.Detail[i].GetValue('PUL_NOMZONE'),False);
    End;
  End;
  Cbox := GetParamSoc('SO_PGDUEFAXRESP');
  SetControlText('EFAXCORRESP',Cbox);

  SMessage := 'DUE créée le '+TobDUE.Detail[13].GetValue('DONNEEAFFICH')+ 'à '+ TobDUE.Detail[14].GetValue('DONNEEAFFICH');
  SetControlText('INFO1',SMessage);
End;

//===========//
Procedure TOF_PGDUESAISIE.DUE_ENREGISTRER_EN_TABLE;
Var
  i : Integer;
  CTRL, Nom, ValDUE, ValSai, Cqualif : String;
  TobDonnee, TFille, TobSuivi : Tob;
  Defaut : THEdit;
  YY,MM,JJ : WORD;

Begin
  TobDonnee := Tob.Create('DUEDONNEE',Nil,-1);
  DecodeDate(Date,YY,MM,JJ);

  For i := 0 To TobDUE.Detail.Count -1 do
  Begin
     TFille := Tob.Create('DUEDONNEE',TobDonnee,-1);

     Nom    := TobDUE.Detail[i].GetValue('PUL_NOMZONE');
     CTRL   := TobDUE.Detail[i].GetValue('PUL_DUEVALEUR');
     ValSai := Trim(TobDUE.Detail[i].GetValue('DONNEEAFFICH'));

     If (Pos('Valeur',CTRL) <> 0) Then
        ValDUE := Copy(CTRL,8,Length(CTRL))
      Else If (Pos('DATE',CTRL) <> 0) Then
      Begin
        If ValSai <> '01/01/1900' Then ForceNumerique(ValSai,ValDUE) Else ValDUE := '';
      End
      Else If (Pos('HEURE',CTRL) <> 0) Then
      Begin
        ForceNumerique(ValSai,ValDUE);
        ValDUE := Copy(ValDUE,1,4);
      End
      Else If (Nom = 'C_TYPE_DECL') Then
        ValDUE := Copy(ValSai,2,1)
      Else If (Nom = 'C_MT_DCL') And (ValSai <> '') Then
        ValDUE := 'MT' + ValSai
      Else If (Nom = 'C_CVLT_SAL') Then
      Begin
        If ValSai = 'MME' Then ValDUE := '2'
           Else If ValSai = 'MLE' Then ValDUE := '3'
           Else ValDUE := '1';
      End
      Else If (Nom = 'C_TYPE_CONTRAT') Then
      Begin
        If ValSai = 'CDD' Then ValDUE := '1'
           Else If ValSai = 'CDI' Then ValDUE := '2'
           Else If ValSai = 'CTT' Then ValDUE := '3'
           Else If ValSai = 'CNE' Then ValDUE := '4'
           Else ValDUE := '';
      End
      Else If (Pos('ECRAN',CTRL) <> 0) Then
      Begin
        Defaut := ThEdit(Getcontrol(Nom));
        If (Defaut <> nil) And (Valsai <> '') Then ValDUE := Defaut.Text Else ValDUE := '';
      End
      Else If (Nom = 'C_EMPLOI') Then
      Begin
        If Length(ValSai)= 2 Then
        Begin
          ValDUE  := Copy(ValSai,1,1);
          Cqualif := Copy(ValSai,2,1);
        End
        Else
        Begin
          ValDUE  := ValSai;
          Cqualif := '';
        End;
      End
      Else If (Nom = 'C_QUALIF') Then
        ValDUE := Cqualif
     Else
      ValDUE := ValSai;

     TFille.PutValue('PUD_SALARIE',Sal_Matr);
     TFille.PutValue('PUD_ANNEE',IntToStr(YY));
     TFille.PutValue('PUD_ORDRE',IntToStr(Nordre));
     TFille.PutValue('PUD_ORDRESEG',IntToStr(i));
     TFille.PutValue('PUD_NOMZONE',TobDUE.Detail[i].GetValue('PUL_NOMZONE'));
     TFille.PutValue('PUD_POSZONE',TobDUE.Detail[i].GetValue('PUL_POSZONE'));
     TFille.PutValue('PUD_DONNEE',ValDUE);
     TFille.PutValue('PUD_DONNEEAFFICH',TobDUE.Detail[i].GetValue('DONNEEAFFICH'));
     TFille.InsertOrUpdateDB(False);
  End;
  //TOBDebug(TobDonnee);
  TobDonnee.Free;
  If Demande = 'CREATION' Then
  Begin
     TobSuivi := Tob.Create('DUESUIVI',Nil,-1);

     TobSuivi.PutValue('PUS_PREDEFINI','DOS');
     TobSuivi.PutValue('PUS_NODOSSIER',V_PGI.Nodossier);
     TobSuivi.PutValue('PUS_SALARIE',Sal_Matr);
     TobSuivi.PutValue('PUS_ANNEE',IntToStr(YY));
     TobSuivi.PutValue('PUS_ORDRE',IntToStr(Nordre));
     TobSuivi.PutValue('PUS_DATECREATION',TobDUE.Detail[13].GetValue('DONNEEAFFICH'));
     If CtrlDUE = True Then TobSuivi.PutValue('PUS_ETATDUE','1')
                       Else TobSuivi.PutValue('PUS_ETATDUE','0');
     TobSuivi.PutValue('PUS_FORMULAIRE','-');
     TobSuivi.InsertOrUpdateDB(False);
     TobSuivi.Free;
  End;
  FaireMaj := False;
End;

//===========//
Initialization
  registerclasses ( [ TOF_PGDUESAISIE ] ) ;
end.
