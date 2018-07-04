{***********UNITE*************************************************
Auteur  ......  :
Créé le ...... : 18/07/2002
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : SAISIEINSCBUDGET ()
Mots clefs ... : TOF;SAISIEINSCBUDGET
*****************************************************************
PT1 24/04/2003  V_42 JL  Développement pour compatibilité CWAS
PT2 16/07/2007  V_72 FC  FQ 14510 Obligé de cliquer deux fois sur la croix rouge pour pouvoir sortie
PT3 10/09/2007 JL V_80 FQ 14706 Prise en compte valeur <<Aucun>> dans historique
PT4 18/12/2007 FC V_81 FQ 15064 Libellé coefficient, indice, qualification, niveau : mal géré quand
                                plusieurs prédéfinis
}
Unit UTofPGSaisielisteHistorique;

Interface

Uses StdCtrls,Controls,Classes,
{$IFNDEF EAGLCLIENT}
     db,     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}Fe_Main,
{$ELSE}
      MaineAgl,
{$ENDIF}
     forms,ParamDat,sysutils,ComCtrls,HCtrls,HEnt1,HMsgBox,UTOB,UTOF,uTableFiltre,SaisieList,
     hTB97,PGOutilsFormation,EntPaie,LookUp;


Type
  TOF_PGSAISIELISTEHISTO = Class (TOF)

    procedure OnLoad                   ; override ;
    procedure OnClose                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    Private
    Niveau,Libelle,Salarie,Champ,Activite,TypeDonnee : String;
    TF :  TTableFiltre;
    SQLLookUp,TabletteLookUp,LibelleLookUp : String; //PT4
    procedure DateElipsisclick(Sender : TObject);
    procedure TauxClick(Sender : TObject);
    procedure OnEnterProfil(Sender: Tobject);
    procedure ChargeProfilActivite(LeChamp, Nat: string);
    procedure ChampElipsisclick(Sender: Tobject);  //PT4
  end ;
var PGTobSalListe,PGTobTableDynListe : Tob;


Implementation

procedure TOF_PGSAISIELISTEHISTO.OnLoad ;
begin
  Inherited ;
  If TF.TypeAction <> TaConsult then
  begin
    If TF.Reccount =0 then
    begin
      TF.Insert;
    end;
    SetFocuscontrol('PHD_NEWVALEUR');
  end;
  if TF.TypeAction <> TaConsult then TFSaisieList(Ecran).Caption := 'Saisie de la zone : '+Libelle
  else TFSaisieList(Ecran).Caption := 'Consultation historique de la zone : '+Libelle;
  UpdateCaption(TFSaisieList(Ecran));
end ;

procedure TOF_PGSAISIELISTEHISTO.OnClose ;
var Valeur : String;
    LaDate,DateMax : TDateTime;
begin
  Inherited ;
  TF.First;
  LaDate := IDate1900;
  DateMax := IDate1900;
  Valeur := '';
  While Not TF.EOF do
  begin
    LaDate := TF.GetValue('PHD_DATEAPPLIC');
    If LaDate <= V_PGI.DateEntree then
    begin
      If LaDate > DateMax Then
      begin
        DateMax := LaDate;
        Valeur := TF.GetValue('PHD_NEWVALEUR');
        if (TF.GetValue('PHD_TRAITEMENTOK') <> 'X') and (TF.GetValue('PHD_NEWVALEUR') <> '') and (TF.GetValue('PHD_NEWVALEUR') <> '-') Then  //PT2
        begin
          TF.PutValue('PHD_TRAITEMENTOK','X');
          TF.Post;
        end;
      end;
    end;
    TF.Next;
  end;
  If((TypeDonnee = 'F') or (TypeDonnee = 'I')) and (Valeur = '') then Valeur := '0';//PT3
   TFSaisieList(Ecran).Retour := Valeur;
  If Assigned(PGTobSalListe) then FreeANdNil(PGTobSalListe);
  If Assigned(PGTobTableDynListe) then FreeANdNil(PGTobTableDynListe);
end;

procedure TOF_PGSAISIELISTEHISTO.OnArgument (S : String ) ;
var Tablette : String;
    Q,Qt : TQuery;
    TypeInfo,CodeTabLib,Theme : String;
    Combo : THValComboBox;
    Edit,TauxAT : THEdit;
    NiveauDefaut,TypePop : String;
    St,Nature,Convention:String;
begin
  Inherited ;
       TF  :=  TFSaisieList(Ecran).LeFiltre;
       Niveau := Trim(ReadTokenPipe(S, ';'));
       Salarie := Trim(ReadTokenPipe(S, ';'));
       Champ := Trim(ReadTokenPipe(S, ';'));
       if Champ = 'PSA_SITUATIONFAMIL' then Champ := 'PSA_SITUATIONFAMI';
       Tablette := Trim(ReadTokenPipe(S, ';'));
       TypePop := Trim(ReadTokenPipe(S, ';'));
       Q := OpenSQL('SELECT * FROM PARAMSALARIE WHERE PPP_PGINFOSMODIF="'+Champ+'"',True);
       If Not Q.Eof then
       begin
        Libelle := Q.FindField('PPP_LIBELLE').AsString;
        TypeInfo := Q.FindField('PPP_PGTYPEINFOLS').AsString;
        NiveauDefaut := Q.FindField('PPP_TYPENIVEAU').AsString;
        TypeDonnee := Q.FindField('PPP_PGTYPEDONNE').AsString;
        Theme := Q.FindField('PPP_PGTHEMESALARIE').AsString;
       end;
       Ferme(Q);
       If Champ = 'PSA_TRAVAILN1' then Libelle := VH_Paie.PGLibelleOrgStat1
       else If Champ = 'PSA_TRAVAILN2' then Libelle := VH_Paie.PGLibelleOrgStat2
       else If Champ = 'PSA_TRAVAILN3' then Libelle := VH_Paie.PGLibelleOrgStat3
       else If Champ = 'PSA_TRAVAILN4' then Libelle := VH_Paie.PGLibelleOrgStat4
       else If Champ = 'PSA_SALAIREMOIS1' then Libelle := VH_Paie.PgSalLib1+' mensuel'
       else If Champ = 'PSA_SALAIREMOIS2' then Libelle := VH_Paie.PgSalLib2+' mensuel'
       else If Champ = 'PSA_SALAIREMOIS3' then Libelle := VH_Paie.PgSalLib3+' mensuel'
       else If Champ = 'PSA_SALAIREMOIS4' then Libelle := VH_Paie.PgSalLib4+' mensuel'
       else If Champ = 'PSA_SALAIREMOIS5' then Libelle := VH_Paie.PgSalLib5+' mensuel'
       else If Champ = 'PSA_SALAIRANN1' then Libelle := VH_Paie.PgSalLib1+' annuel'
       else If Champ = 'PSA_SALAIRANN2' then Libelle := VH_Paie.PgSalLib2+' annuel'
       else If Champ = 'PSA_SALAIRANN3' then Libelle := VH_Paie.PgSalLib3+' annuel'
       else If Champ = 'PSA_SALAIRANN4' then Libelle := VH_Paie.PgSalLib4+' annuel'
       else If Champ = 'PSA_SALAIRANN5' then Libelle := VH_Paie.PgSalLib5+' annuel';
       SetControlText('NIVSAISIE',Niveau);
       SetControlEnabled('NIVSAISIE',False);
       SetControlcaption('NIVPREC','Niveau préconisé : '+RechDom('PGNIVEAUAVDOS',NiveauDefaut,False));
       If TypeInfo = 'ZLS' then TFSaisieList(Ecran).Caption := 'Elément complémentaire salarié : '+Libelle
       else
       begin
        SetControlVisible('NIVSAISIE',False);
        SetControlVisible('TNIVSAISIE',False);
        SetControlVisible('NIVPREC',False);

       end;
              If Theme = 'PRO' then
       begin
        Q := OpenSQL('SELECT PSA_ACTIVITE FROM SALARIES WHERE PSA_SALARIE="'+Salarie+'"',True);
        If Not Q.eof then Activite := Q.FindField('PSA_ACTIVITE').AsString
        else Activite := '';
        Ferme(Q);
        Combo := THValComboBox(getcontrol('PHD_NEWVALEUR'));
        if Combo <> nil then Combo.OnEnter := OnEnterProfil;
       end
       else Activite := '';
       If (Tablette <> '') or ( champ = 'PSA_ORDREAT' ) then
       begin
        If champ <> 'PSA_ORDREAT' then SetControlProperty('PHD_NEWVALEUR','DataType',Tablette)
        else
        begin
          TauxAT := THEdit(GetControl('PHD_NEWVALEUR'));
          if TauxAT <> nil then TauxAt.OnElipsisClick := TauxClick;
        end;
        If TFSaisieList(Ecran).Name = 'PGLISTEHISTODE' then SetControlProperty('PHD_NEWVALEUR','ElipsisButton',True);
       end;
       If Niveau = 'SAL' then TF.WhereTable := 'WHERE PHD_SALARIE="'+Salarie+'" AND PHD_PGINFOSMODIF="'+Champ+'"'
       else If Niveau = 'POP' then TF.WhereTable := 'WHERE PHD_POPULATION="'+TypePop+'" AND PHD_CODEPOP="'+Salarie+'" AND PHD_PGINFOSMODIF="'+Champ+'"'
       else If Niveau = 'ETB' then TF.WhereTable := 'WHERE PHD_ETABLISSEMENT="'+Salarie+'" AND PHD_PGINFOSMODIF="'+Champ+'"'
       else TF.WhereTable := 'WHERE PHD_SALARIE="" AND  PHD_ETABLISSEMENT="" AND PHD_CODEPOP="" AND PHD_POPULATION="" AND PHD_PGINFOSMODIF="'+Champ+'"';
       TFSaisieList(Ecran).Retour := '';
       If Tablette = 'PGCOMBOZONELIBRE' then
       begin
        Q := OpenSQL('SELECT PPP_CODTABL FROM PARAMSALARIE WHERE PPP_PGINFOSMODIF="'+Champ+'"',True);
        If Not Q.Eof then
        begin
          CodeTabLib := Q.FindField('PPP_CODTABL').AsString;
        end;
        Ferme(Q);
        Q := OpenSQL('SELECT PSA_SALARIE,PSA_CONVENTION,PSA_ETABLISSEMENT FROM SALARIES WHERE PSA_SALARIE="'+Salarie+'"',True);
        PGTobSalListe := Tob.Create('LeSalarie',Nil,-1);
        PGTobSalListe.LoadDetailDB('LeSalarie','','',Q,False);
        Ferme(Q);
        Q := OpenSQL('SELECT * FROM TABLEDIMDET WHERE PTD_CODTABL="'+CodeTabLib+'" ORDER BY PTD_DTVALID DESC',True);
        PGTobTableDynListe := Tob.Create('ParamTableDyn',Nil,-1);
        PGTobTableDynListe.LoadDetailDB('ParamTableDyn','','',Q,False);
        Ferme(Q);
       end;
       If TypeDonnee = 'D' then
       begin
        Edit := THEdit(GetControl('PHD_NEWVALEUR'));
        If edit <> Nil then
        begin
          Edit.ElipsisButton := True;
          Edit.OnElipsisClick := DateElipsisclick;
        end;
       end;
   If TypeInfo = 'SAL' then
   begin
     If TypeDonnee = 'D' then TF.LaGridListe := 'PGHISTODSAISIEDAT'
     else If TypeDonnee = 'T' then TF.LaGridListe := 'PGHISTODSAISIECOM'
     else If TypeDonnee = 'B' then TF.LaGridListe := 'PGHISTODSAISIECH'
     else TF.LaGridListe := 'PGHISTODSAISIE';
   end
   else
   begin
     If TypeDonnee = 'D' then TF.LaGridListe := 'PGHISTODSAISIEDAT'
     else If TypeDonnee = 'T' then TF.LaGridListe := 'PGHISTODSAISIEZLC'
     else If TypeDonnee = 'B' then TF.LaGridListe := 'PGHISTODSAISIECH'
     else TF.LaGridListe := 'PGHISTODSAISIEZL';
   end;
   //DEB PT4
   If (Tablette = 'PGLIBCOEFFICIENT') or (Tablette = 'PGLIBINDICE') or (Tablette = 'PGLIBNIVEAU') or (Tablette = 'PGLIBQUALIFICATION') then
   begin
      Edit := THEdit(GetControl('PHD_NEWVALEUR'));
      If Edit <> Nil then
        Edit.OnElipsisClick := ChampElipsisclick;
      TabletteLookUp := Tablette;
      LibelleLookUp := Libelle;
   end;
   SQLLookUp := '';
   Nature := '';
   Convention := '';
   If (Champ = 'PSA_QUALIFICATION') or (Champ = 'PSA_COEFFICIENT') or (Champ = 'PSA_INDICE') or (Champ = 'PSA_NIVEAU') then
   begin
    if (Champ = 'PSA_QUALIFICATION') then
      Nature := 'QUA';
    if (Champ = 'PSA_COEFFICIENT') then
      Nature := 'COE';
    if (Champ = 'PSA_INDICE') then
      Nature := 'IND';
    if (Champ = 'PSA_NIVEAU') then
      Nature := 'NIV';
    Q:= OpenSQL('SELECT PSA_CONVENTION FROM SALARIES WHERE PSA_SALARIE="'+Salarie+'"',True);
    If Not Q.Eof then
      Convention := Q.FindField('PSA_CONVENTION').AsString;
    Ferme(Q);
    SQLLookUp := 'Select PMI_CODE,PMI_LIBELLE FROM MINIMUMCONVENT M1' +
      ' WHERE  ##PMI_PREDEFINI## PMI_NATURE="' + Nature + '"' +
      ' AND (PMI_CONVENTION="'+Convention+'" OR PMI_CONVENTION="000")' +
      ' AND (PMI_PREDEFINI="DOS"' +
      ' OR (PMI_PREDEFINI="STD" AND PMI_CONVENTION="'+Convention+'" ' +
      '   AND NOT EXISTS(SELECT PMI_CODE FROM MINIMUMCONVENT ' +
      '   WHERE ##PMI_PREDEFINI## PMI_CODE=M1.PMI_CODE AND PMI_PREDEFINI="DOS"))' +
      ' OR (PMI_PREDEFINI="STD" AND PMI_CONVENTION="000" ' +
      '   AND NOT EXISTS(SELECT PMI_CODE FROM MINIMUMCONVENT '+
      '   WHERE ##PMI_PREDEFINI## PMI_CODE=M1.PMI_CODE ' +
      '   AND (PMI_PREDEFINI="DOS" OR (PMI_PREDEFINI="STD" AND PMI_CONVENTION="'+Convention+'")))))';
//      SetControlProperty('PHD_NEWVALEUR','Plus',' AND (PMI_CONVENTION="'+Q.FindField('PSA_CONVENTION').AsString+'"'+
//      ' OR PMI_CONVENTION="000")');
{   If Champ = 'PSA_COEFFICIENT' then
   begin
    Q:= OpenSQL('SELECT PSA_CONVENTION FROM SALARIES WHERE PSA_SALARIE="'+Salarie+'"',True);
    If Not Q.Eof then SetControlProperty('PHD_NEWVALEUR','Plus',' AND (PMI_CONVENTION="'+Q.FindField('PSA_CONVENTION').AsString+'" OR'+
                         ' PMI_CONVENTION="000")');
    Ferme(Q);
    end;
    If Champ = 'PSA_INDICE' then
   begin
    Q:= OpenSQL('SELECT PSA_CONVENTION FROM SALARIES WHERE PSA_SALARIE="'+Salarie+'"',True);
    If Not Q.Eof then SetControlProperty('PHD_NEWVALEUR','Plus',' AND (PMI_CONVENTION="'+Q.FindField('PSA_CONVENTION').AsString+'" OR'+
                         ' PMI_CONVENTION="000")');
    Ferme(Q);
    end;
     If Champ = 'PSA_NIVEAU' then
   begin
    Q:= OpenSQL('SELECT PSA_CONVENTION FROM SALARIES WHERE PSA_SALARIE="'+Salarie+'"',True);
    If Not Q.Eof then SetControlProperty('PHD_NEWVALEUR','Plus',' AND (PMI_CONVENTION="'+Q.FindField('PSA_CONVENTION').AsString+'" OR'+
                         ' PMI_CONVENTION="000")');
    Ferme(Q);
    end;
}
//FIN PT4
  end;
end ;

procedure TOF_PGSAISIELISTEHISTO.DateElipsisclick(Sender : TObject);
var key : char;
begin
    key  :=  '*';
    ParamDate (Ecran, Sender, Key);
end;

procedure TOF_PGSAISIELISTEHISTO.TauxClick(Sender: TObject);
var
  StSql: string;
  Edit : THEdit;
  Etab : String;
  Q : TQuery;
begin
  Q := OpenSQL('SELECT PSA_ETABLISSEMENT FROM SALARIES WHERE PSA_SALARIE="'+Salarie+'"',True);
  Etab := Q.FindField('PSA_ETABLISSEMENT').AsString;
  Ferme(Q);
  if Etab = '' then exit;
  StSql := ' PAT_ETABLISSEMENT ="' + Etab + '"';
  Edit := THEdit(GetControl('PHD_NEWVALEUR'));
  if Edit <> nil then LookUpList(Edit, 'Taux AT', 'TAUXAT', 'PAT_ORDREAT', 'PAT_LIBELLE', StSql, 'PAT_ORDREAT', TRUE, -1);
end;

procedure TOF_PGSAISIELISTEHISTO.OnEnterProfil(Sender: Tobject);
var
  Name: string;
begin
  if sender = nil then exit;
  { Affectation de la propriété plus pour chaque profil gérant l'activité }
  if Champ = 'PSA_PROFILAFP' then ChargeProfilActivite('PSA_PROFILAFP', 'AFP')
  else if Champ = 'PSA_PROFILANCIEN' then ChargeProfilActivite('PSA_PROFILANCIEN', 'ANC')
  else if Champ = 'PSA_PROFILAPP' then ChargeProfilActivite('PSA_PROFILAPP', 'APP')
  else if Champ = 'PSA_PROFILCDD' then ChargeProfilActivite('PSA_PROFILCDD', 'CDD')
  else if Champ = 'PSA_PROFILCGE' then ChargeProfilActivite('PSA_PROFILCGE', 'CGE')
  else if Champ = 'PSA_PROFILFNAL' then ChargeProfilActivite('PSA_PROFILFNAL', 'FNL')
  else if Champ = 'PSA_PROFILMUT' then ChargeProfilActivite('PSA_PROFILMUT', 'MUT')
  else if Champ = 'PSA_PERIODBUL' then ChargeProfilActivite('PSA_PERIODBUL', 'PER')
  else if Champ = 'PSA_PROFILPRE' then ChargeProfilActivite('PSA_PROFILPRE', 'PRE')
  else if Champ = 'PSA_PROFIL' then ChargeProfilActivite('PSA_PROFIL', 'PRO')
  else if Champ = 'PSA_PROFILRBS' then ChargeProfilActivite('PSA_PROFILRBS', 'RBS')
  else if Champ = 'PSA_PROFILREM' then ChargeProfilActivite('PSA_PROFILREM', 'REM')
  else if Champ = 'PSA_PROFILRET' then ChargeProfilActivite('PSA_PROFILRET', 'RET')
  else if Champ = 'PSA_REDREPAS' then ChargeProfilActivite('PSA_REDREPAS', 'RRE')
  else if Champ = 'PSA_REDRTT1' then ChargeProfilActivite('PSA_REDRTT1', 'RT1')
  else if Champ = 'PSA_REDRTT2' then ChargeProfilActivite('PSA_REDRTT2', 'RT2')
  else if Champ = 'PSA_PROFILTPS' then ChargeProfilActivite('PSA_PROFILTPS', 'TPS')
  else if Champ = 'PSA_PROFILTRANS' then ChargeProfilActivite('PSA_PROFILTRANS', 'TRA')
  else if Champ = 'PSA_PROFILTSS' then ChargeProfilActivite('PSA_PROFILTSS', 'TSS');
end;

procedure TOF_PGSAISIELISTEHISTO.ChargeProfilActivite(LeChamp, Nat: string);
begin
  if RechDom('PGTYPEPROFIL', Nat, True) = 'ACT' then
  begin
    if GetField('PSA_ACTIVITE') <> '' then
      SetControlProperty('PHD_NEWVALEUR', 'Plus', 'AND (PPI_ACTIVITE="' +  Activite + '" OR PPI_ACTIVITE="" OR PPI_ACTIVITE IS NULL)')
    else
      SetControlProperty('PHD_NEWVALEUR', 'Plus', '');
  end;
end;

//DEB PT4
procedure TOF_PGSAISIELISTEHISTO.ChampElipsisclick(Sender: Tobject);
var
  Edit : THEdit;
begin
  Edit := THEdit(GetControl('PHD_NEWVALEUR'));
  If Edit <> Nil then
    LookUpList(Edit, LibelleLookUp, TabletteLookUp, 'PMI_CODE', 'PMI_LIBELLE', '', 'PMI_CODE', TRUE, -1,SQLLookUp);
end;
//FIN PT4

Initialization
  registerclasses ( [ TOF_PGSAISIELISTEHISTO ] ) ;
end.

