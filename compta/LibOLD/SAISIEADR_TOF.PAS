{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 06/03/2007
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : SAISIEADR ()
Mots clefs ... : TOF;SAISIEADR
*****************************************************************}
Unit SAISIEADR_TOF ;

Interface

Uses Adresse, StdCtrls,
     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     FE_Main, // AGLLanceFiche
     mul,
{$else}
     MaineAGL, // AGLLanceFiche
     eMul,
{$ENDIF} 
     uTob,
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox, 
     UTOF;
// Lancement de la fiche AGL
procedure YY_LanceSaisieAdresse;

Type
  TOF_SAISIEADR = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  private
    // Pour la saisie :
    Pays          : string;
    TypePays      : string;
    TypeAdr       : string;
    SSTypeAdr     : integer;
    DateSais      : TDateTime;
    Adresse       : TAdresse;
    Destinataire  : string;
    KeepValue     : boolean;
    // Pour la restitution :
    PaysRest          : string;
    TypePaysRest      : string;
    TypeAdrRest       : string;
    SSTypeAdrRest     : integer;
    DateSaisRest      : TDateTime;
    DestinataireRest  : string;
    // Evenements fiche
    procedure BValiderClick(Sender : TObject); 
    procedure BModifierClick(Sender : TObject);   
    procedure BSupprimerClick(Sender : TObject);
    procedure BGetViaTOBClick(Sender : TObject);
    procedure BGetViaMEMOClick(Sender : TObject);   
    procedure BImprimerClick(Sender : TObject);
    procedure GetGuidClick(Sender : TObject);  
    procedure GetDateClick(Sender : TObject);
    procedure Create500000enreg(Sender : TObject);
    procedure CreateTiers(Sender : TObject);      
    procedure CreateEnreg(Sender : TObject);
    // Evenements Saisie :        
    procedure PaysChange(Sender : TObject);      
    procedure AdresseChange(Sender : TObject);
    procedure TypeAdresseChange(Sender : TObject);
    procedure SSTypeAdresseChange(Sender : TObject);  
    procedure DestinataireChange(Sender : TObject);
    procedure DestDispoChange(Sender : TObject);
    procedure DateSaisieChange(Sender : TObject);   
    procedure HeureSaisieChange(Sender : TObject);  
    procedure KeepValueChange(Sender : TObject);
    // Evenements Restitution :
    procedure TypeRestChange(Sender : TObject);  
    procedure PaysRestChange(Sender : TObject);
    procedure AdresseRestChange(Sender : TObject);
    procedure TypeAdresseRestChange(Sender : TObject);
    procedure SSTypeAdresseRestChange(Sender : TObject);    
    procedure DestinataireRestChange(Sender : TObject);
    procedure DateSaisieRestChange(Sender : TObject);   
    procedure HeureSaisieRestChange(Sender : TObject);
    // Communs :
    procedure ChargeCombo;
  end ;

Implementation
uses StrUtils
{$IFDEF EAGLCLIENT}
     , UtileAGL
{$ELSE}
     , EdtREtat
{$ENDIF}
     , utilPGI
     , HTB97
     , ParamSoc
     , uLibCFONB
     , uTobDebug
     , uLibExportTVA
     , uLibGestionAdresse ;

procedure YY_LanceSaisieAdresse;
begin
  AGLLanceFiche('YY', 'SAISIEADR', '', '', '');
end;

//////////////////////////////////////////////////////////////////////////////
///             FONCTIONS POUR LA GESTION DE LA FICHE
//////////////////////////////////////////////////////////////////////////////


procedure TOF_SAISIEADR.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_SAISIEADR.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_SAISIEADR.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_SAISIEADR.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_SAISIEADR.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_SAISIEADR.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_SAISIEADR.OnArgument (S : String ) ;
begin
  Inherited ;
  // Evenements :         
  TToolBarButton97(GetControl('GETMEMO'          )).OnClick  := BGetViaMEMOClick; 
  TToolBarButton97(GetControl('GETTOB'           )).OnClick  := BGetViaTOBClick; 
  TToolbarButton97(GetControl('BSUPPRIMER'       )).OnClick  := BSupprimerClick; 
  TToolbarButton97(GetControl('CREATEENREG'      )).OnClick  := CreateEnreg;
  TToolBarButton97(GetControl('BMODIFIER'        )).OnClick  := BModifierClick;
  TToolbarButton97(GetControl('BImprimer'        )).OnClick  := BImprimerClick;
  TToolBarButton97(GetControl('BValider'         )).OnClick  := BValiderClick;
  TToolbarButton97(GetControl('GETGUID'          )).OnClick  := GetGuidClick;
  TToolbarButton97(GetControl('GETDATE'          )).OnClick  := GetDateClick;
  THValComboBox(   GetControl('SSTYPEADRESSEREST')).OnChange := SSTypeAdresseRestChange;
  THValComboBox(   GetControl('TYPEADRESSEREST'  )).OnChange := TypeAdresseRestChange;
  THValComboBox(   GetControl('SSTYPEADRESSE'    )).OnChange := SSTypeAdresseChange;
  THValComboBox(   GetControl('TYPEADRESSE'      )).OnChange := TypeAdresseChange;
  THValComboBox(   GetControl('ADRESSEREST'      )).OnChange := AdresseRestChange;
  THValComboBox(   GetControl('DESTDISPO'        )).OnChange := DestDispoChange;
  THValComboBox(   GetControl('PAYSREST'         )).OnChange := PaysRestChange;
  THValComboBox(   GetControl('ADRESSE'          )).OnChange := AdresseChange;
  THValComboBox(   GetControl('PAYS'             )).OnChange := PaysChange;
  THRadioButton(   GetControl('RTYPERESTLIG'     )).OnClick  := TypeRestChange; 
  THRadioButton(   GetControl('RTYPEREST'        )).OnClick  := TypeRestChange;
  THCheckBox(      GetControl('KEEPVALUE'        )).OnClick  := KeepValueChange;
  THEdit(          GetControl('DATEREST'         )).OnExit   := DateSaisieRestChange;
  THEdit(          GetControl('HEUREREST'        )).OnExit   := HeureSaisieRestChange;
  THEdit(          GetControl('DESTINATAIRE'     )).OnChange := DestinataireChange;
  THEdit(          GetControl('DATE'             )).OnExit   := DateSaisieChange;
  THEdit(          GetControl('HEURE'            )).OnExit   := HeureSaisieChange;




  // Renseignement de la combo
  ChargeCombo;

  // Chargement des valeurs               
  THEdit(GetControl('DATEREST')).EditMask := '';
  THEdit(GetControl('DATE')).EditMask := '';
  THValComboBox(GetControl('TYPEADRESSE')).Itemindex := 0;
  THValComboBox(GetControl('DESTDISPO')).Itemindex := 0;
  THValComboBox(GetControl('PAYS')).Itemindex := THValComboBox(GetControl('PAYS')).Values.IndexOf(GetParamSocSecur('SO_PAYS',''));
  SetControlText('DATE',FormatDateTime('dd/mm/yyyy',now));
  SetControlText('HEURE',FormatDateTime('hh:mm',now));
  THValComboBox(GetControl('TYPEADRESSEREST')).Itemindex := 0;
  //THValComboBox(GetControl('DESTDISPOREST')).Itemindex := 0;
  THValComboBox(GetControl('PAYSREST')).Itemindex := THValComboBox(GetControl('PAYSREST')).Values.IndexOf(GetParamSocSecur('SO_PAYS',''));
  SetControlText('DATEREST',FormatDateTime('dd/mm/yyyy',now));
  SetControlText('HEUREREST',FormatDateTime('hh:mm',now));

  // Récuperation des valeurs
  if Trim(THValComboBox(GetControl('DESTDISPO')).Value) <> '' then
     SetControlText('DESTINATAIRE',Trim(THValComboBox(GetControl('DESTDISPO')).Value))
  else
     GetGuidClick(nil);
  TypeAdresseChange(nil);  
  SSTypeAdresseChange(nil);
  DestinataireChange(nil);
  PaysChange(nil);
  TypeAdresseRestChange(nil);
  SSTypeAdresseRestChange(nil);
  DestinataireRestChange(nil);
  PaysRestChange(nil);
  TypeRestChange(nil);
  DateSaisieChange(nil); 
end ;

procedure TOF_SAISIEADR.OnClose ;
begin
  Inherited ;
  // On libere le composant si existe
  if assigned(Adresse) then
     FreeAndNil(Adresse);
end ;

//////////////////////////////////////////////////////////////////////////////
///             FONCTIONS POUR LA GESTION DES BOUTONS
//////////////////////////////////////////////////////////////////////////////

procedure TOF_SAISIEADR.BValiderClick(Sender : TObject);
begin    
  if not(assigned(Adresse)) then Exit;
  // On sauvegarde les données adresse :
  Adresse.SaveData;
  // On recharge la combo GUID pour prendre en compte le possible nouveau GUID
  ChargeCombo;
end;

procedure TOF_SAISIEADR.BModifierClick(Sender : TObject);
begin
  if not(assigned(Adresse)) then Exit;
  // On modifie les données adresse :
  Adresse.SaveData(true);
  // On recharge la combo GUID pour prendre en compte le possible nouveau GUID
  ChargeCombo;
end;

procedure TOF_SAISIEADR.BSupprimerClick(Sender : TObject);
begin
  if not assigned(Adresse) then exit;
  // On supprime l'adresse courante :
  Adresse.SupprimerAdresse;
  if Adresse.LastError <> 0 then
     PGIError(Adresse.LastErrorMsg)
  else
     PGIInfo('Suppression effectuée');
end;

procedure TOF_SAISIEADR.BGetViaTOBClick(Sender : TObject);
var           
  monMEMO : TMemo;
  maTOB : TOB;
  maFille : TOB;
begin                     
  if not(assigned(Adresse)) then Exit;
  // On force la sauvegarde pour que les données de la table soient les mêmes que sur la fiche
  BValiderClick(nil);
  monMEMO := nil;
  maTOB := TOB.Create('Une TOB',nil,-1);
  maFille :=  TOB.Create('',maTOB,-1);
  try
     // Récuperation des données
     if GetControlText('RTYPERESTLIG') = 'X' then
        RestitutionAdresse(maFille,monMEMO,-1,PaysRest,StrToInt(TypePaysRest),GetControlText('DESTDISPOREST'),TypeAdrRest,SSTypeAdrRest,DateSaisRest,StrToInt(GetControlText('NB_LIGNE')))
     else
        RestitutionAdresseType(maFille,monMEMO,-1,PaysRest,StrToInt(TypePaysRest),GetControlText('DESTDISPOREST'),TypeAdrRest,SSTypeAdrRest,DateSaisRest,GetControlText('TYPELIGNEREST'));
     TobDebug(maTOB);
  finally
     FreeAndNil(maTOB);
  end;
end;

procedure TOF_SAISIEADR.BGetViaMEMOClick(Sender : TObject);
var
  monMEMO : TMemo;
  maTOB : TOB;
begin           
  if not(assigned(Adresse)) then Exit;
  // On force la sauvegarde pour que les données de la table soient les mêmes que sur la fiche
  BValiderClick(nil);
  monMEMO := TMemo(GetControl('MEMO2'));
  maTOB := nil;
  monMEMO.Clear;
  // Récuperation des données
  if GetControlText('RTYPERESTLIG') = 'X' then
     RestitutionAdresse(maTOB,monMEMO,-1,PaysRest,StrToInt(TypePaysRest),GetControlText('DESTDISPOREST'),TypeAdrRest,SSTypeAdrRest,DateSaisRest,StrToInt(GetControlText('NB_LIGNE')))
  else
     RestitutionAdresseType(maTOB,monMEMO,-1,PaysRest,StrToInt(TypePaysRest),GetControlText('DESTDISPOREST'),TypeAdrRest,SSTypeAdrRest,DateSaisRest,GetControlText('TYPELIGNEREST'));
end;

procedure TOF_SAISIEADR.BImprimerClick(Sender : TObject);
var           
  monMEMO : TMemo;
  maTOB   : TOB;
  maFille : TOB;
begin         
  if not(assigned(Adresse)) then Exit;
  BValiderClick(nil);
  monMEMO := nil;
  maTOB   := TOB.Create('Une TOB',nil,-1);
  // Pour l'etat il faut renseigné la fille
  maFille :=  TOB.Create('',maTOB,-1);
  try
     // Récuperation des données
     if GetControlText('RTYPERESTLIG') = 'X' then
        RestitutionAdresse(maFille,monMEMO,-1,PaysRest,StrToInt(TypePaysRest),GetControlText('DESTDISPOREST'),TypeAdrRest,SSTypeAdrRest,DateSaisRest,StrToInt(GetControlText('NB_LIGNE')))
     else
        RestitutionAdresseType(maFille,monMEMO,-1,PaysRest,StrToInt(TypePaysRest),GetControlText('DESTDISPOREST'),TypeAdrRest,SSTypeAdrRest,DateSaisRest,GetControlText('TYPELIGNEREST'));
     // Lancement de l'Etat
     LanceEtatTOB('E','INT','TST',maTOB,true,false,false,nil,'',TraduireMemoire('Etat de Test'),false);
  finally
     FreeAndNil(maTOB);
  end;

end;

//////////////////////////////////////////////////////////////////////////////
///             FONCTIONS POUR LA SAISIE
//////////////////////////////////////////////////////////////////////////////

procedure TOF_SAISIEADR.GetGuidClick(Sender : TObject);
begin
  // Permet de créer un GUID
  SetControlText('DESTINATAIRE', AglGetGuid)
end;

procedure TOF_SAISIEADR.GetDateClick(Sender : TObject);
var
  date : tDateTime;
begin
  // Permet de créer un GUID
  if not(assigned(Adresse)) then Exit;
  date := Adresse.GetDateSaisie;
  if date > 0 then
  SetControlText('DATE', FormatDateTime('dd/mm/yyyy',date)); 
  SetControlText('HEURE', FormatDateTime('hh:mm',date));
end;

procedure TOF_SAISIEADR.Create500000enreg(Sender : TObject);
var
  maTOB : TOB;
  fille : TOB;
  i,j   : integer;
begin
for j:=1 to 50 do
begin
  maTOB := TOB.Create('VALEURADRE',nil,-1);
  try
     for i := 1 to 10000 do
     begin
        fille := TOB.Create('VALEURADRE',maTOB,-1);
        fille.PutValue('VAD_PARCODE',1);
        fille.PutValue('VAD_DEST',AglGetGuid);
        fille.PutValue('VAD_TYPEADR','LIV;');
        fille.PutValue('VAD_SSTYPEADR',0);
        fille.PutValue('VAD_DATESAISIE',now);
        fille.PutValue('VAD_CHAMP_1','NOM');
        fille.PutValue('VAD_CHAMP_2','PRENOM');
        fille.PutValue('VAD_CHAMP_3','ADRESSE 1');
        fille.PutValue('VAD_CHAMP_4','ADRESSE 2');
        fille.PutValue('VAD_CHAMP_5','ADRESSE 3');
        fille.PutValue('VAD_CHAMP_6','69000');
        fille.PutValue('VAD_CHAMP_7','LYON');
        fille.PutValue('VAD_CHAMP_8','PAYS');
     end;
     maTOB.InsertOrUpdateDB(true); 
  finally
     maTOB.Free;
  end;          
end;
end;

procedure TOF_SAISIEADR.CreateEnreg(Sender : TObject);
var
  Guid  : string;
  monAdresse : TOB;
  filleAdresse : TOB;
  maTOB : TOB;
  mafille : TOB;
  i,j   : integer;
begin
  for j := 0 to 1 do
  begin
     maTOB := TOB.Create('TIERS',nil,-1);
     monAdresse := TOB.Create('VALEURADRE',nil,-1);
     try
        GUID := AglGetGuid;
        // Création Tiers
        maTOB.PutValue('T_GUID',GUID);
        maTOB.PutValue('T_AUXILIAIRE',CompleteWithCar(IntToStr(1 + (j * 10000) + 610000),'0',8,'G'));
        maTOB.PutValue('T_LIBELLE','LIBELLE');
        maTOB.PutValue('T_COLLECTIF','41110000');
        maTOB.PutValue('T_ADRESSE1','ADRESSE 1');
        maTOB.PutValue('T_CODEPOSTAL','69000');
        maTOB.PutValue('T_VILLE','LYON');
        maTOB.PutValue('T_SECTEUR','...');
        maTOB.PutValue('T_MODEREGLE','001');
        maTOB.PutValue('T_REGIMETVA','FRA');
        maTOB.PutValue('T_LETTRABLE','X');
        maTOB.PutValue('T_DATECREATION',now);
        maTOB.PutValue('T_DATEMODIF',now);
        maTOB.PutValue('T_SOLDEPROGRESSIF','X');
        maTOB.PutValue('T_SAUTPAGE','X');
        maTOB.PutValue('T_TOTAUXMENSUELS','X');
        maTOB.PutValue('T_SOCIETE','001');
        // Création adresse corespondante
        monAdresse.PutValue('VAD_PARCODE',2);
        monAdresse.PutValue('VAD_DEST',GUID);
        monAdresse.PutValue('VAD_TYPEADR','LIV;');
        monAdresse.PutValue('VAD_SSTYPEADR',0);
        monAdresse.PutValue('VAD_DATESAISIE',now);
        monAdresse.PutValue('VAD_CHAMP_1','NOM');
        monAdresse.PutValue('VAD_CHAMP_2','PRENOM');
        monAdresse.PutValue('VAD_CHAMP_3','ADRESSE 1');
        monAdresse.PutValue('VAD_CHAMP_4','ADRESSE 2');
        monAdresse.PutValue('VAD_CHAMP_5','ADRESSE 3');
        monAdresse.PutValue('VAD_CHAMP_6','69000');
        monAdresse.PutValue('VAD_CHAMP_7','LYON');
        monAdresse.PutValue('VAD_CHAMP_8','PAYS');
        for i := 2 to 10000 do
        begin                   
           GUID := AglGetGuid;
           // Création Tiers;
           maFille := TOB.Create('TIERS',maTOB,-1);
           maFille.PutValue('T_GUID',GUID);
           maFille.PutValue('T_AUXILIAIRE',CompleteWithCar(IntToStr(i + (j * 10000) + 610000),'0',8,'G'));
           maFille.PutValue('T_LIBELLE','LIBELLE');
           maFille.PutValue('T_COLLECTIF','41110000');
           maFille.PutValue('T_ADRESSE1','ADRESSE 1');
           maFille.PutValue('T_CODEPOSTAL','69000');
           maFille.PutValue('T_VILLE','LYON');
           maFille.PutValue('T_SECTEUR','...');
           maFille.PutValue('T_MODEREGLE','001');
           maFille.PutValue('T_REGIMETVA','FRA');
           maFille.PutValue('T_LETTRABLE','X');
           maFille.PutValue('T_DATECREATION',now);
           maFille.PutValue('T_DATEMODIF',now);
           maFille.PutValue('T_SOLDEPROGRESSIF','X');
           maFille.PutValue('T_SAUTPAGE','X');
           maFille.PutValue('T_TOTAUXMENSUELS','X');
           maFille.PutValue('T_SOCIETE','001');
           // Création Adresse
           filleAdresse := TOB.Create('VALEURADRE',monAdresse,-1);
           filleAdresse.PutValue('VAD_PARCODE',2);
           filleAdresse.PutValue('VAD_DEST',GUID);
           filleAdresse.PutValue('VAD_TYPEADR','LIV;');
           filleAdresse.PutValue('VAD_SSTYPEADR',0);
           filleAdresse.PutValue('VAD_DATESAISIE',now);
           filleAdresse.PutValue('VAD_CHAMP_1','NOM');
           filleAdresse.PutValue('VAD_CHAMP_2','PRENOM');
           filleAdresse.PutValue('VAD_CHAMP_3','ADRESSE 1');
           filleAdresse.PutValue('VAD_CHAMP_4','ADRESSE 2');
           filleAdresse.PutValue('VAD_CHAMP_5','ADRESSE 3');
           filleAdresse.PutValue('VAD_CHAMP_6','69000');
           filleAdresse.PutValue('VAD_CHAMP_7','LYON');
           filleAdresse.PutValue('VAD_CHAMP_8','PAYS');
        end;
        maTOB.InsertDB(nil);
        monAdresse.InsertDB(nil);
     finally
        maTOB.Free;
        monAdresse.Free;
     end;
  end;
end;

procedure TOF_SAISIEADR.CreateTiers(Sender : TObject);
var
  SQL   : string;
  Guid  : TOB;
  maTOB : TOB;
  mafille : TOB;
  i,j   : integer;
  Q     : TQuery;
begin
  SQL := 'SELECT VAD_DEST FROM VALEURADRE';
  Q := OpenSQL(SQL,true);
  GUID := TOB.Create('ma tob',nil,-1);
  GUID.LoadDetailDB('VALEURADRE','','',Q,false);
  for j := 0 to 200 do
  begin
     maTOB := TOB.Create('TIERS',nil,-1);
     try
        maTOB := TOB.Create('TIERS',maTOB,-1);
        maTOB.PutValue('T_GUID',GUID.Detail[1 + (j * 10000)].GetValue('VAD_DEST'));
        maTOB.PutValue('T_AUXILIAIRE',CompleteWithCar(IntToStr(1 + (j * 10000)),'0',8,'G'));
        maTOB.PutValue('T_LIBELLE','LIBELLE');
        maTOB.PutValue('T_COLLECTIF','41110000');
        maTOB.PutValue('T_ADRESSE1','ADRESSE 1');
        maTOB.PutValue('T_CODEPOSTAL','69000');
        maTOB.PutValue('T_VILLE','LYON');
        maTOB.PutValue('T_SECTEUR','...');
        maTOB.PutValue('T_MODEREGLE','001');
        maTOB.PutValue('T_REGIMETVA','FRA');
        maTOB.PutValue('T_LETTRABLE','X');
        maTOB.PutValue('T_DATECREATION',now);
        maTOB.PutValue('T_DATEMODIF',now);
        maTOB.PutValue('T_SOLDEPROGRESSIF','X');
        maTOB.PutValue('T_SAUTPAGE','X');
        maTOB.PutValue('T_TOTAUXMENSUELS','X');
        maTOB.PutValue('T_SOCIETE','001');
        for i := 2 to 10000 do
        begin
           maFille := TOB.Create('TIERS',maTOB,-1);
           maFille.PutValue('T_GUID',GUID.Detail[i + (j * 10000)].GetValue('VAD_DEST'));
           maFille.PutValue('T_AUXILIAIRE',CompleteWithCar(IntToStr(i + (j * 10000)),'0',8,'G'));
           maFille.PutValue('T_LIBELLE','LIBELLE');
           maFille.PutValue('T_COLLECTIF','41110000');
           maFille.PutValue('T_ADRESSE1','ADRESSE 1');
           maFille.PutValue('T_CODEPOSTAL','69000');
           maFille.PutValue('T_VILLE','LYON');
           maFille.PutValue('T_SECTEUR','...');
           maFille.PutValue('T_MODEREGLE','001');
           maFille.PutValue('T_REGIMETVA','FRA');
           maFille.PutValue('T_LETTRABLE','X');
           maFille.PutValue('T_DATECREATION',now);
           maFille.PutValue('T_DATEMODIF',now);
           maFille.PutValue('T_SOLDEPROGRESSIF','X');
           maFille.PutValue('T_SAUTPAGE','X');
           maFille.PutValue('T_TOTAUXMENSUELS','X');
           maFille.PutValue('T_SOCIETE','001');
        end;
        maTOB.InsertDB(nil);
     finally
        maTOB.Free;
     end;
  end;
  Ferme(Q);
  GUID.Free;
end;

procedure TOF_SAISIEADR.PaysChange(Sender : TObject);
begin
  // Le pays n'a pas changé
  if codeISOduPays(Trim(GetControlText('PAYS'))) = Pays then Exit;
  Pays := codeISOduPays(Trim(GetControlText('PAYS')));
  // On met à jour l'ellipsis du type pays
  THValComboBox(GetControl('ADRESSE')).Plus := ' PAD_PAYS = "' + Pays + '"';
  // On selectionne le premier element de la liste par defaut
  THValComboBox(GetControl('ADRESSE')).ItemIndex := 0;
  // On met à jour le changement
  AdresseChange(nil);
end;

procedure TOF_SAISIEADR.AdresseChange(Sender : TObject);
begin
  // Le type n'a pas changé
  if TypePays = Trim(GetControlText('ADRESSE')) then Exit;
  TypePays := Trim(GetControlText('ADRESSE'));
  // Si pas de type selectionné on passe -1 par defaut
  if not(isNumeric(TypePays)) then TypePays := '-1';
  // Si l'objet existe dejà on le libere
  if assigned(Adresse) then FreeAndNil(Adresse);
  // On crée le composant Adresse avec les bons paramètres :
  Adresse := TAdresse.Create(ecran,Destinataire,Pays,StrToInt(TypePays),TypeAdr,SSTypeAdr,DateSais);
  if Adresse.LastError <> 0 then
  begin
     PGIError(Adresse.LastErrorMsg);
     Exit;
  end;
  // On affecte le memo
  Adresse.Affichage := TMemo(GetControl('MEMO'));
  // On met à jour l'affichage
  Adresse.UpdateAffichage;
end;

procedure TOF_SAISIEADR.TypeAdresseChange(Sender : TObject);
begin
  // Le type n'a pas changé
  if TypeAdr = Trim(GetControlText('TYPEADRESSE')) then Exit;
  if THMultiValComboBox(GetControl('TYPEADRESSE')).Tous then
     TypeAdr := 'AFA;INT;LIV;PIE;TIE;'
  else if THMultiValComboBox(GetControl('TYPEADRESSE')).Aucun then
     TypeAdr := ''
  else
     TypeAdr := Trim(GetControlText('TYPEADRESSE'));
  // Si le composant Adresse existe on le recharge en changeant uniquement le type adresse
  // car le pavé adresse reste le même.
  if assigned(Adresse) then
     Adresse.ChangeTypeAdresse(TypeAdr,SSTypeAdr,KeepValue);
end;

procedure TOF_SAISIEADR.SSTypeAdresseChange(Sender : TObject);
begin
  // Le type n'a pas changé
  if SSTypeAdr = StrToInt(GetControlText('SSTYPEADRESSE')) then Exit;
  SSTypeAdr := StrToInt(Trim(GetControlText('SSTYPEADRESSE')));
  // Si le composant Adresse existe on le recharge en changeant uniquement le type adresse
  // car le pavé adresse reste le même.
  if assigned(Adresse) then
     Adresse.ChangeTypeAdresse(TypeAdr,SSTypeAdr,KeepValue);
end;


procedure TOF_SAISIEADR.DestinataireChange(Sender : TObject);
begin
  // Le destinataire n'a pas changé
  if GetControlText('DESTINATAIRE') = Destinataire then Exit;
  Destinataire := GetControlText('DESTINATAIRE');
  // Si le composant existe on change uniquement le destinataire
  if assigned(Adresse) then
     Adresse.ChangeDestinataire(Destinataire,TypeAdr,SSTypeAdr,DateSais,KeepValue);
end;

procedure TOF_SAISIEADR.DestDispoChange(Sender : TObject);
begin
  // On met à jour le THEdit on fonction de l'element selectionné dans la combo
  SetControlText('DESTINATAIRE',GetControlText('DESTDISPO'));
end;

procedure TOF_SAISIEADR.DateSaisieChange(Sender : TObject);
begin
  if Trim(GetControlText('DATE')) = '0' then DateSais := 0
  else if Trim(GetControlText('DATE')) = '' then DateSais := 0
  else if Trim(GetControlText('DATE')) = '//' then DateSais := 0
  else if Trim(GetControlText('HEURE')) = '0' then DateSais := 0
  else if Trim(GetControlText('HEURE')) = '' then DateSais := 0
  else if Trim(GetControlText('HEURE')) = ':' then DateSais := 0
  else
     // On met à jour le THEdit on fonction de l'element selectionné dans la combo
     DateSais :=  StrToDateTime(GetControlText('DATE')) + StrToDateTime(GetControlText('HEURE'));
  // Si le composant existe on change uniquement le destinataire
  if assigned(Adresse) then
     Adresse.ChangeDateSaisie(DateSais,KeepValue);
end;

procedure TOF_SAISIEADR.HeureSaisieChange(Sender : TObject);
begin
  if Trim(GetControlText('DATE')) = '0' then DateSais := 0
  else if Trim(GetControlText('DATE')) = '' then DateSais := 0
  else if Trim(GetControlText('DATE')) = '//' then DateSais := 0
  else if Trim(GetControlText('HEURE')) = '0' then DateSais := 0
  else if Trim(GetControlText('HEURE')) = '' then DateSais := 0
  else if Trim(GetControlText('HEURE')) = ':' then DateSais := 0
  else
     // On met à jour le THEdit on fonction de l'element selectionné dans la combo
     DateSais :=  StrToDateTime(GetControlText('DATE')) + StrToDateTime(GetControlText('HEURE'));
  // Si le composant existe on change uniquement le destinataire
  if assigned(Adresse) then
     Adresse.ChangeDateSaisie(DateSais,KeepValue);
end;


procedure TOF_SAISIEADR.KeepValueChange(Sender : TObject);
begin
  if GetCheckBoxState('KEEPVALUE') = cbUnchecked then
     KeepValue := false
  else if GetCheckBoxState('KEEPVALUE') = cbChecked then
     KeepValue := true;
end;

//////////////////////////////////////////////////////////////////////////////
///             FONCTIONS POUR LA RESTITUTION
//////////////////////////////////////////////////////////////////////////////

procedure TOF_SAISIEADR.TypeRestChange(Sender : TObject);
begin
  if GetControlText('RTYPERESTLIG') = 'X' then
  begin
     SetControlEnabled('NB_LIGNE',true);
     SetControlEnabled('TYPELIGNEREST',false);
  end
  else
  begin
     SetControlEnabled('NB_LIGNE',false);
     SetControlEnabled('TYPELIGNEREST',true);
     THValComboBox(GetControl('TYPELIGNEREST')).ItemIndex := 0;
  end;
end;

procedure TOF_SAISIEADR.PaysRestChange(Sender : TObject);
begin
  // Le pays n'a pas changé
  if codeISOduPays(Trim(GetControlText('PAYSREST'))) = PaysRest then Exit;
  PaysRest := codeISOduPays(Trim(GetControlText('PAYSREST')));
  // On met à jour l'ellipsis du type pays
  THValComboBox(GetControl('ADRESSEREST')).Plus := ' PAD_PAYS = "' + PaysRest + '"';
  // On selectionne le premier element de la liste par defaut
  THValComboBox(GetControl('ADRESSEREST')).ItemIndex := 0;
  // On met à jour le changement
  AdresseRestChange(nil);
end;

procedure TOF_SAISIEADR.AdresseRestChange(Sender : TObject);
begin
  // Le type n'a pas changé
  if TypePaysRest = Trim(GetControlText('ADRESSEREST')) then Exit;
  TypePaysRest := Trim(GetControlText('ADRESSEREST'));
  // Si pas de type selectionné on passe -1 par defaut
  if not(isNumeric(TypePaysRest)) then TypePaysRest := '-1';
end;

procedure TOF_SAISIEADR.TypeAdresseRestChange(Sender : TObject);
begin
  // Le type n'a pas changé
  if TypeAdrRest = Trim(GetControlText('TYPEADRESSEREST')) then Exit;
  TypeAdrRest := Trim(GetControlText('TYPEADRESSEREST'));
end;

procedure TOF_SAISIEADR.SSTypeAdresseRestChange(Sender : TObject);
begin
  // Le type n'a pas changé
  if SSTypeAdrRest = StrToInt(GetControlText('SSTYPEADRESSEREST')) then Exit;
  SSTypeAdrRest := StrToInt(Trim(GetControlText('SSTYPEADRESSEREST')));
end;


procedure TOF_SAISIEADR.DestinataireRestChange(Sender : TObject);
begin
  // Le destinataire n'a pas changé
  if GetControlText('DESTDISPOREST') = DestinataireRest then Exit;
  DestinataireRest := GetControlText('DESTDISPOREST');
end;

procedure TOF_SAISIEADR.DateSaisieRestChange(Sender : TObject);
begin
  if Trim(GetControlText('DATEREST')) = '0' then DateSais := 0
  else if Trim(GetControlText('DATEREST')) = '' then DateSais := 0
  else if Trim(GetControlText('DATEREST')) = '//' then DateSais := 0
  else if Trim(GetControlText('HEUREREST')) = '0' then DateSais := 0
  else if Trim(GetControlText('HEUREREST')) = '' then DateSais := 0
  else if Trim(GetControlText('HEUREREST')) = ':' then DateSais := 0
  else
     // On met à jour le THEdit on fonction de l'element selectionné dans la combo
     DateSaisRest :=  StrToDateTime(GetControlText('DATEREST')) + StrToDateTime(GetControlText('HEUREREST'));
end;
procedure TOF_SAISIEADR.HeureSaisieRestChange(Sender : TObject);
begin
  if Trim(GetControlText('DATEREST')) = '0' then DateSais := 0
  else if Trim(GetControlText('DATEREST')) = '' then DateSais := 0
  else if Trim(GetControlText('DATEREST')) = '//' then DateSais := 0
  else if Trim(GetControlText('HEUREREST')) = '0' then DateSais := 0
  else if Trim(GetControlText('HEUREREST')) = '' then DateSais := 0
  else if Trim(GetControlText('HEUREREST')) = ':' then DateSais := 0
  else
     // On met à jour le THEdit on fonction de l'element selectionné dans la combo
     DateSaisRest :=  StrToDateTime(GetControlText('DATEREST')) + StrToDateTime(GetControlText('HEUREREST'));
end;
//////////////////////////////////////////////////////////////////////////////
///             FONCTIONS COMMUNES
//////////////////////////////////////////////////////////////////////////////

procedure TOF_SAISIEADR.ChargeCombo;
var
  SQL : string;
  Combo : THValComboBox;
begin
  exit;
  // Charge le THValComboBox avec les GUID Distinct de la table VALEURADRE
  SQL := 'SELECT VAD_DEST,VAD_DEST FROM VALEURADRE GROUP BY VAD_DEST';
  Combo := THValComboBox(GetControl('DESTDISPO'));
  ChargeTHValComboBox(SQL,Combo);                       
  Combo := THValComboBox(GetControl('DESTDISPOREST'));
  ChargeTHValComboBox(SQL,Combo);
end;

Initialization
  registerclasses ( [ TOF_SAISIEADR ] ) ;
end.
