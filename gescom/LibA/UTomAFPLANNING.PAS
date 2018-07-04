{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 26/03/2002
Modifié le ... :   /  /
Description .. : Source TOM de la TABLE : AFPLANNING (AFPLANNING)
Mots clefs ... : TOM;AFPLANNING
*****************************************************************}
Unit UTomAFPLANNING;

Interface
      
uses {$IFDEF VER150} variants,{$ENDIF} Graphics,
     StdCtrls,
     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     db, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} FE_Main,Fiche, FichList,dbctrls,
{$ELSE}
     MaineAGL, eFiche,  eFichList,spin,
{$ENDIF}

{$IFDEF BTP}
	 CalcOleGenericBTP,
{$ENDIF}
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     UTOM,
      AglInit, HDB,
     UTob, HTB97, HPanel, Vierge, paramsoc,

     //UtilPlanning,
     UtilGC, FactUtil, dicobtp,
     AFUtilArticle, UtilTaches, AFPlanningCst, AffaireUtil,
     UtilRessource, uafo_ressource;

Type
  TOM_AFPLANNING = Class (TOM)
    procedure OnNewRecord                ; override ;
    procedure OnDeleteRecord             ; override ;
    procedure OnUpdateRecord             ; override ;
    procedure OnAfterUpdateRecord        ; override ;
    procedure OnLoadRecord               ; override ;
    procedure OnChangeField ( F: TField) ; override ;
    procedure OnArgument ( S: String )   ; override ;
    procedure OnClose                    ; override ;
    procedure OnCancelRecord             ; override ;

  Private

    fCB_APL_UNITETEMPS  : THDBValComboBox;
    fBT_Ferme           : TToolBarButton97;
    fBT_Valider         : TToolBarButton97;
    fED_DATEDEBPLA      : THDBEdit;
    fED_DATEFINPLA      : THDBEdit;
    fED_RESSOURCE       : THDBEdit;
    fED_QTEPLANIFIEE    : THDBEdit;

    fED_DATEDEBREAL     : THDBEdit;
    fED_DATEFINREAL     : THDBEdit;
    fED_QTEREALISE      : THDBEdit;

    fFO_PLANNING        : TForm;
    fED_ATA_NUMEROTACHE : THedit;

    fMaTob              : Tob;
    //fPA_HEUREPLA        : THPanel;
    //fPA_HEUREREAL       : THPanel;

    fStAffaire          : String;
    fStFonction         : String;
    fStDateDeb          : String;
    fStDateFin          : String;
    fStNumeroTache      : String;
    fStRetour           : String;
    fStAction           : String;
    fStPlusAffaire      : String;
    fStSQL              : String;
    fBModiflot          : Boolean;

    procedure CreateControls;
    procedure APL_QTEPLANIFIEEOnChange(Sender: TObject);
    procedure APL_QTEREALISEOnChange(Sender: TObject);
    procedure APL_UNITETEMPSOnChange(Sender: TObject);

    procedure bFermeClick(Sender: TObject);
    procedure bValiderClick(Sender: TObject);
    procedure DATEDEBPLAExit(Sender: TObject);
    procedure DATEFINPLAExit(Sender: TObject);
    procedure DATEDEBREALExit(Sender: TObject);
    procedure DATEFINREALExit(Sender: TObject);
    procedure QTEPLANIFIEEExit(Sender: TObject);
    procedure QTEREALISEExit(Sender: TObject);
    procedure bMemoOnClick(SEnder: TObject);

    procedure ATANumeroTacheOnChange(Sender: TObject);
    procedure APL_RESSOURCEOnChange(Sender: TObject);
    procedure LoadRecord;
    Function PlusAffaire : string;
    procedure UpdateMontant;
    procedure SetArguments(StSQL : string);
end;

Function AFLanceFicheAFPLANNING(pStCle, pStParam : String) : String;

const
	// libellés des messages
	TexteMessage: array[1..8] of string 	= (
          {1}        'Le libellé doit être renseigné.',
          {2}        'L''état doit être renseigné.',
          {3}        'La tâche doit être renseignée.',
          {4}        'Cette ressource n''est pas prévue pour cette tâche.',
          {5}        'La date de fin doit être supérieure ou égale à la date de début.',
          {6}        'La date de début doit être inférieure ou égale à la date de fin.',
          {7}        'La ressource doit être renseignée.',
          {8}        'Ce numéro de tâche n''existe pas pour cette affaire et cette ressource.'
                     );

Implementation

function AFLanceFicheAFPLANNING(pStCle, pStParam : String) : String;
begin
  result := AGLLanceFiche('AFF','AFPLANNING','', pStCle,pStParam);
end;

procedure TOM_AFPLANNING.OnNewRecord ;
var
  vStAff0,vStAff1,vStAff2 : String;
  vStAff3,vStAvenant      : String;

begin

  LoadRecord;
  Inherited;

  TPageControl(GetControl('Pages')).ActivePage := TTabSheet(GetControl('PGeneral'));

  SetField('APL_TYPELIGNEPLA', 'TAC');
  SetField('APL_ETATLIGNE', getparamsoc('SO_AFPLANNINGETAT'));

  // C.B 28/01/03
  SetField('APL_DATEMODIFPLA', now);
  SetField('APL_DATEDEBPLA', strToDate(fStDateDeb));
  SetField('APL_DATEFINPLA', strToDate(fStDateFin));
  SetField('APL_DATEDEBREAL', strToDate(fStDateDeb));
  SetField('APL_DATEFINREAL', strToDate(fStDateFin));
   
  SetField('APL_AFFAIRE', fStAffaire);
  SetField('APL_FONCTION', fStFonction);
  SetField('APL_TIERS', fMaTob.Getvalue('APL_TIERS'));

  fED_RESSOURCE.OnChange := nil;
  SetField('APL_RESSOURCE', fMaTob.Getvalue('APL_RESSOURCE'));
  fED_RESSOURCE.OnChange := APL_RESSOURCEOnChange;

  SetField('APL_LIBELLEPLA', fMaTob.Getvalue('APL_LIBELLEPLA'));
  fED_ATA_NUMEROTACHE.Text := fStNumeroTache;

  // pour l'instant, on initialise avec jour comme unité de saisie
  setField('APL_UNITETEMPS', 'J');
   
  if getparamsoc('SO_AFPLANDECHARGE') then setControlEnabled('APL_UNITETEMPS', False);
  SetField('APL_QTEPLANIFIEE', StrToDate(fStDateFin) - StrToDate(fStDateDeb) + 1);
//  SetField('APL_QTEREALISE', StrToDate(fStDateFin) - StrToDate(fStDateDeb) + 1);

  {$IFDEF BTP}
  BTPCodeAffaireDecoupe(fStAffaire,vStAff0,vStAff1,vStAff2,vStAff3,vStAvenant, taModif, false);
  {$ELSE}
  CodeAffaireDecoupe(fStAffaire,vStAff0,vStAff1,vStAff2,vStAff3,vStAvenant, taModif, false);
  {$ENDIF}

  SetField('APL_AFFAIRE0',vStAff0);
  SetField('APL_AFFAIRE1',vStAff1);
  SetField('APL_AFFAIRE2',vStAff2);
  SetField('APL_AFFAIRE3',vStAff3);
  SetField('APL_AVENANT',vStAvenant);

  // heures pas gérées
  SetField('APL_HEUREDEBPLA', '9');
  SetField('APL_HEUREFINPLA', '18');
  SetField('APL_HEUREDEBREAL', '9');
  SetField('APL_HEUREFINREAL', '18');

end;

procedure TOM_AFPLANNING.OnDeleteRecord ;
begin
  Inherited ;
end ;

procedure TOM_AFPLANNING.OnUpdateRecord;
var
  vTOBAffaires        : Tob;
  vTOBArticles        : Tob;
  vAFOAssistants      : TAFO_Ressources;
  //stringlist_libelle  : TStrings;
  //stringlist_valeurs  : TStrings;
  //vQr                 : TQuery;
  //vTob                : Tob;
  //vTobTache           : Tob;
                            
begin
  Inherited;

  If (GetField ('APL_LIBELLEPLA') = '') then
    begin
      SetFocusControl ('APL_LIBELLEPLA');
      LastError := 1;
      LastErrorMsg := TexteMessage[LastError] ;
    end

  else If (GetField ('APL_ETATLIGNE') = '') then
    begin
      SetFocusControl ('APL_ETATLIGNE');
      LastError := 2;
      LastErrorMsg := TexteMessage[LastError] ;
    end

  else If (GetControlText ('ATA_NUMEROTACHE') = '') or
          (GetControlText ('ATA_NUMEROTACHE') = '0') then
    begin
      SetFocusControl ('ATA_NUMEROTACHE');
      LastError := 3;
      LastErrorMsg := TexteMessage[LastError] ;
    end

  // la ressource est obligatoire
  else If (GetControlText ('APL_RESSOURCE') = '')  then
    begin
      SetFocusControl ('APL_RESSOURCE');
      LastError := 7;
      LastErrorMsg := TexteMessage[LastError] ;
    end

  // la tache n'existe pas   ??? AB - ('ATA_FAMILLETACHE')???
  else if (GetControlText ('ATA_FAMILLETACHE') = '') then
    begin
      SetFocusControl ('ATA_NUMEROTACHE');
      LastError := 8;
      LastErrorMsg := TexteMessage[LastError] ;
    end

  else
    begin
      if (varAsType(GetField('APL_NUMEROLIGNE'),VarString) = '') or
         (varAsType(GetField('APL_NUMEROLIGNE'),VarInteger) = 0) then
        begin
//Modif FV   SetField('APL_NUMEROLIGNE',GetNumLignePlanning(GetField('APL_AFFAIRE')));

          if GetParamSoc('SO_AFALIGNREALISE') then
            begin
              SetField('APL_QTEREALISE', GetField('APL_QTEPLANIFIEE'));
              SetField('APL_QTEREALUREF', GetField('APL_QTEPLANIFUREF'));
            end;
          // on aligne les dates de realisation sur les dates de planification
          SetField('APL_DATEDEBREAL', GetField('APL_DATEDEBPLA'));
          SetField('APL_DATEFINREAL', GetField('APL_DATEFINPLA'));
        end;

      fMaTob.PutValue('APL_QTEPLANIFIEE', GetControlText('APL_QTEPLANIFIEE'));
      fMaTob.PutValue('APL_QTEREALISE', GetControlText('APL_QTEREALISE'));
      fMaTob.PutValue('APL_QTEPLANIFUREF', GetControlText('APL_QTEPLANIFUREF'));
      fMaTob.PutValue('APL_QTEREALUREF', GetControlText('APL_QTEREALUREF'));
                                             
      vAFOAssistants  := TAFO_Ressources.Create;
      vTOBAffaires    := TOB.Create('Les Affaires',Nil,-1);
      vTOBArticles    := TOB.Create('les Articles',Nil,-1) ;
      try                                         
//Modif FV Valorisation(fMaTob, 'APL', vAFOAssistants, vTOBAffaires, vTobArticles);
      finally
        vTOBAffaires.Free;
        vTobArticles.Free;
        vAFOAssistants.Free;
      end;

      UpdateMontant;
    end;
end;

procedure TOM_AFPLANNING.OnAfterUpdateRecord ;
begin
  Inherited ;
end ;

procedure TOM_AFPLANNING.OnLoadRecord ;
begin

  if fMaTob.Getvalue('APL_RESSOURCE') <> '' then
    SetControlEnabled('APL_RESSOURCE', false)
  else
    SetControlEnabled('APL_RESSOURCE', true);

  // filtre des taches sur les affaires
  fED_ATA_NUMEROTACHE.plus := PlusAffaire;

  // C.B 28/01/03
  if GetField('APL_NUMEROTACHE') <> '0' then
    SetControlText('ATA_NUMEROTACHE',GetField('APL_NUMEROTACHE'));

  if fBModifLot then SetArguments(fStSQL);
  ATANumeroTacheOnChange(self);
  loadRecord;
                         
  // on est en consultation si modification interdite
{Modif FV  if not EtatModifAutorisee(GetControlText('APL_ETATLIGNE')) then
    begin
      TFfiche(Ecran).typeaction := TaConsult;
      self.ModifAutorisee:=false;
      setControlEnabled ('BValider', false);
      SetControlText('LA_ETAT', 'L''état ' + rechdom('AFTETAT', GetControlText('APL_ETATLIGNE'), false) + ' ne permet pas de modification');
      SetControlVisible('LA_ETAT', true);
    end;}

  Ecran.Caption := TraduitGa('Planning ' +  GetField('APL_LIBELLEPLA') + ' : ');
//  UpdateCaption(ecran);

  Inherited;
end;

procedure TOM_AFPLANNING.OnChangeField ( F: TField ) ;
begin
  Inherited;
end;

procedure TOM_AFPLANNING.OnArgument ( S: String );
var
  Champ   : String;
  Valeur  : String;
  Tmp     : String;
  Bt      : TToolbarButton97;
  x       : integer;
  CleUnique : string;
  CleTob  : string;  
begin

  Inherited;
  x := pos('MODIFLOT',S);
  fBModiflot := x<>0;
  if fBModiflot then
  begin
    fMaTob:=Tob.Create('AFPLANNING', nil, -1);
    CleUnique := TFfiche(Ecran).FLequel ;
    CleTob := '"'+ReadTokenSt (CleUnique)+'";';
    CleTob := CleTob + ReadTokenSt (CleUnique)+';';
    fMaTob.SelectDB(CleTob,nil);
    TFfiche(Ecran).MonoFiche:=true;
    fStSQL := copy(S,x+9,length(S));
    fStNumeroTache  := fMaTob.Getvalue('APL_NUMEROTACHE');
    SetControlEnabled('ATA_NUMEROTACHE', false);
    SetControlEnabled('APL_DATEDEBPLA', false);
    SetControlEnabled('APL_DATEFINPLA', false);
    SetControlEnabled('APL_HEUREDEBPLA', false);
    SetControlEnabled('APL_HEUREFINPLA', false);
    fStAction := 'MODIFICATION';
  end else
    fMaTob := theTob;

  fStAffaire  := fMaTob.Getvalue('APL_AFFAIRE');
  fStFonction := fMaTob.Getvalue('APL_FONCTION');
  fStDateDeb  := fMaTob.Getvalue('APL_DATEDEBPLA');
  fStDateFin  := fMaTob.Getvalue('APL_DATEFINPLA');

  CreateControls;

  fED_ATA_NUMEROTACHE.OnChange := ATANumeroTacheOnChange;
  SetControlText('ATA_NUMEROTACHE', fStNumeroTache);
  //  ATANumeroTacheOnChange(self);

  // traitement des arguments
  Tmp:=(Trim(ReadTokenSt(S)));
  While (Tmp <>'') do
    Begin
      DecodeArgument(Tmp, Champ, valeur);
      If Champ='ACTION' then
        begin
          If valeur='MODIFICATION' then
            begin
              fStNumeroTache  := fMaTob.Getvalue('APL_NUMEROTACHE');
              SetControlEnabled('ATA_NUMEROTACHE', false);
              SetControlEnabled('APL_DATEDEBPLA', false);
              SetControlEnabled('APL_DATEFINPLA', false);
              SetControlEnabled('APL_HEUREDEBPLA', false);
              SetControlEnabled('APL_HEUREFINPLA', false);
              fStAction := 'MODIFICATION';
            end

          else If valeur='CREATION' then
            begin
              fStNumeroTache  := '';
              SetControlEnabled('ATA_NUMEROTACHE', true);
              fStAction := 'CREATION';
            end;
        end

      else if Champ = 'DUPLIQUER' then
        begin
          fStNumeroTache  := fMaTob.Getvalue('APL_NUMEROTACHE');
          SetControlEnabled('ATA_NUMEROTACHE', false);
        end;

      Tmp:=(Trim(ReadTokenSt(S)));
    End;

  fBT_Ferme.OnClick := bFermeClick ;
  fBT_Valider.OnClick := bValiderClick ;

  // traduction champs libres
  GCMAJChampLibre (TForm (Ecran), False, 'COMBO', 'APL_LIBRETACHE', 10, '_');
  GCMAJChampLibre (TForm (Ecran), False, 'EDIT', 'APL_DATELIBRE', 3, '_');
  GCMAJChampLibre (TForm (Ecran), False, 'EDIT', 'APL_CHARLIBRE', 3, '_');
  GCMAJChampLibre (TForm (Ecran), False, 'EDIT', 'APL_VALLIBRE', 3, '_');
  GCMAJChampLibre (TForm (Ecran), False, 'BOOL', 'APL_BOOLLIBRE', 3, '_');

  Bt := TToolbarButton97(GetControl('BMEMO'));
  If (Bt<>Nil) then Bt.Onclick:=bMemoOnClick;

  // gestion des fonctions
  SetControlVisible('APL_FONCTION', GetParamSoc('SO_AFFONCTION'));
  SetControlVisible('TAPL_FONCTION', GetParamSoc('SO_AFFONCTION'));

  //
  SetControlVisible('bfirst', false);
  SetControlVisible('bnext', false);
  SetControlVisible('bprev', false);
  SetControlVisible('bLast', false);

  setControlVisible('PA_HEUREPLA', false);
end;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 25/11/2002
Modifié le ... :
Description .. : reaffecte la tob pour le retour
Mots clefs ... :
*****************************************************************}
procedure TOM_AFPLANNING.OnClose ;
begin
  if fBModifLot then  fMaTob.free else
  if fStRetour <> 'CANCEL' then
    begin
      // theTob a été réinitialisé entre temps
      theTob := fMaTob;
      theTob.PutValue('APL_NUMEROTACHE', GetField('APL_NUMEROTACHE'));
      theTob.PutValue('APL_DATEDEBPLA', GetField('APL_DATEDEBPLA'));
      theTob.PutValue('APL_DATEFINPLA', GetField('APL_DATEFINPLA'));
      theTob.PutValue('APL_LIBELLEPLA', GetField('APL_LIBELLEPLA'));
      theTob.PutValue('APL_RESSOURCE', GetControlText('APL_RESSOURCE'));
      theTob.PutValue('APL_NUMEROLIGNE', GetControlText('APL_NUMEROLIGNE'));
      TFFiche(Ecran).Retour := GetControlText('ATA_LIBELLETACHE1');
    end
  else
    TFFiche(Ecran).Retour := fStRetour;

  Inherited;
end ;

procedure TOM_AFPLANNING.bValiderClick(Sender: TObject);
begin
  inherited;
  fStRetour := 'OK';
  Ecran.Close;
end;

procedure TOM_AFPLANNING.bFermeClick(Sender: TObject);
begin
  inherited;
  fStRetour := 'CANCEL';
  Ecran.Close;
end;

procedure TOM_AFPLANNING.DATEDEBPLAExit(Sender: TObject);
begin
  if IsValidDate(fED_DATEDEBPLA.text) and
     (strtodate(fED_DATEFINPLA.text) < strtodate(fED_DATEDEBPLA.text)) then
    begin
      LastError := 6;
      LastErrorMsg := TexteMessage[LastError];
      PGIBoxAF(TexteMessage[LastError], '');
      fED_DATEDEBPLA.text := fED_DATEFINPLA.text;
    end;
end;

procedure TOM_AFPLANNING.DATEFINPLAExit(Sender: TObject);
begin
  if IsValidDate(fED_DATEFINPLA.text) and
     (strtodate(fED_DATEFINPLA.text) < strtodate(fED_DATEDEBPLA.text)) then
    begin
      LastError := 5;
      LastErrorMsg := TexteMessage[LastError];
      PGIBoxAF(TexteMessage[LastError], '');
      fED_DATEFINPLA.text := fED_DATEDEBPLA.text;
    end;
end;

procedure TOM_AFPLANNING.DATEDEBREALExit(Sender: TObject);
begin
  if IsValidDate(fED_DATEDEBREAL.text) and
     (strtodate(fED_DATEFINREAL.text) < strtodate(fED_DATEDEBREAL.text)) then
    begin
      LastError := 6;
      LastErrorMsg := TexteMessage[LastError];
      PGIBoxAF(TexteMessage[LastError], '');
      fED_DATEDEBREAL.text := fED_DATEFINREAL.text;
    end;
end;

procedure TOM_AFPLANNING.DATEFINREALExit(Sender: TObject);
begin
  if IsValidDate(fED_DATEFINREAL.text) and
     (strtodate(fED_DATEFINREAL.text) < strtodate(fED_DATEDEBREAL.text)) then
    begin
      LastError := 5;
      LastErrorMsg := TexteMessage[LastError];
      PGIBoxAF(TexteMessage[LastError], '');
      fED_DATEFINREAL.text := fED_DATEDEBREAL.text;
    end;
end;

procedure TOM_AFPLANNING.QTEPLANIFIEEExit(Sender: TObject);
begin
  if StrToFloat(fED_QTEPLANIFIEE.Text) < 0 then  fED_QTEPLANIFIEE.Text := '0'
end;

procedure TOM_AFPLANNING.QTEREALISEExit(Sender: TObject);
begin
  if StrToFloat(fED_QTEREALISE.Text) < 0 then fED_QTEREALISE.text := '0'
end;

// controle de l'existance du numero de tâche saisi
procedure TOM_AFPLANNING.ATANumeroTacheOnChange(Sender: TObject);
var
  vSt     : String;
  vQr     : TQuery;
//  vQrArt  : Tquery;

begin

  if (fED_ATA_NUMEROTACHE.Text <> '') and (fED_ATA_NUMEROTACHE.Text <> '0') then
  begin
    vSt := 'SELECT ATA_FAMILLETACHE, ATA_FONCTION, ATA_TIERS, ';
    vSt := vSt + ' ATA_LIBELLETACHE1, ATA_TYPEARTICLE,ATA_ACTIVITEREPRIS, ';
    vSt := vSt + ' ATA_ARTICLE, ATA_CODEARTICLE, ATA_UNITETEMPS ';
    vSt := vSt + ' FROM TACHE ';
    vSt := vSt + ' WHERE ' + copy(fStPlusAffaire,6, length(fStPlusAffaire));
    vSt := vSt + ' AND ATA_NUMEROTACHE =' + fED_ATA_NUMEROTACHE.Text;
    vQr := nil;
    try
      vQR := OpenSql(vSt, True);
      if Not vQR.Eof then
        begin
          SetControlText('ATA_FAMILLETACHE', vQR.Findfield('ATA_FAMILLETACHE').AsString);
          SetControlText('ATA_LIBELLETACHE1', vQR.Findfield('ATA_LIBELLETACHE1').AsString);

          SetField('APL_FONCTION', vQR.Findfield('ATA_FONCTION').AsString);
          SetField('APL_TIERS', vQR.Findfield('ATA_TIERS').AsString);
          SetField('APL_TYPEARTICLE', vQR.Findfield('ATA_TYPEARTICLE').AsString);
          SetField('APL_CODEARTICLE', vQR.Findfield('ATA_CODEARTICLE').AsString);
          SetField('APL_ARTICLE', vQR.Findfield('ATA_ARTICLE').AsString);
          SetField('APL_UNITETEMPS', vQR.Findfield('ATA_UNITETEMPS').AsString);
          if (GetField('APL_NUMEROTACHE') <> fED_ATA_NUMEROTACHE.Text) then  //AB-12/03/03
          SetField('APL_ACTIVITEREPRIS', vQR.Findfield('ATA_ACTIVITEREPRIS').AsString);

          // pour la valorisation
          fMaTob.PutValue('APL_ARTICLE', vQR.Findfield('ATA_ARTICLE').AsString);
          fMaTob.PutValue('APL_TYPEARTICLE', vQR.Findfield('ATA_TYPEARTICLE').AsString);
          fMaTob.PutValue('APL_CODEARTICLE', vQR.Findfield('ATA_CODEARTICLE').AsString);
          fMaTob.PutValue('APL_FONCTION', vQR.Findfield('ATA_FONCTION').AsString);
          fMaTob.PutValue('APL_TIERS', vQR.Findfield('ATA_TIERS').AsString);
          fMaTob.PutValue('APL_UNITETEMPS', vQR.Findfield('ATA_UNITETEMPS').AsString);

          if GetControlText('APL_LIBELLEPLA') = '' then
            SetField('APL_LIBELLEPLA', vQR.Findfield('ATA_LIBELLETACHE1').AsString);
        end;
      SetField('APL_NUMEROTACHE', fED_ATA_NUMEROTACHE.Text);
    finally
      ferme(vQR);
    end;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 05/08/2002
Modifié le ... :   /  /
Description .. : Verifier si la ressource choisie est dans la liste
Description .. : des ressources existant dans cette tache
Mots clefs ... :
*****************************************************************}
procedure TOM_AFPLANNING.APL_RESSOURCEOnChange(Sender: TObject);
begin
  if (GetControlText('APL_RESSOURCE') <> '') then
    begin
      if not ExisteTacheRessource(fStAffaire, fStNumeroTache, GetControlText('APL_RESSOURCE')) then
        begin
          LastError := 4;
          LastErrorMsg := TexteMessage[LastError];
          PGIBoxAF(TexteMessage[LastError], '');
          SetControlText('APL_RESSOURCE', '');
          fED_RESSOURCE.OnChange := nil;
          SetField('APL_RESSOURCE', '');
          fED_RESSOURCE.OnChange := APL_RESSOURCEOnChange;
        end
      else
        SetField('APL_RESSOURCE', '');
    end;
end;     

procedure TOM_AFPLANNING.OnCancelRecord ;
begin                                
  Inherited ;
end ;

procedure TOM_AFPLANNING.APL_QTEPLANIFIEEOnChange(Sender: TObject);
begin
  Inherited;
  SetField('APL_QTEPLANIFUREF', ConversionUnite(fCB_APL_UNITETEMPS.Value, getparamsoc('SO_AFMESUREACTIVITE'), valeur(GetControlText('APL_QTEPLANIFIEE'))));
end;

procedure TOM_AFPLANNING.APL_QTEREALISEOnChange(Sender: TObject);
begin
  Inherited;
  SetField('APL_QTEREALUREF', ConversionUnite(fCB_APL_UNITETEMPS.Value, getparamsoc('SO_AFMESUREACTIVITE'), valeur(GetControlText('APL_QTEREALISE'))));
end;

procedure TOM_AFPLANNING.APL_UNITETEMPSOnChange(Sender: TObject);
begin
  Inherited;
  SetField('APL_QTEPLANIFUREF', ConversionUnite(fCB_APL_UNITETEMPS.Value, getparamsoc('SO_AFMESUREACTIVITE'), valeur(GetControlText('APL_QTEPLANIFIEE'))));
  SetField('APL_QTEREALUREF', ConversionUnite(fCB_APL_UNITETEMPS.Value, getparamsoc('SO_AFMESUREACTIVITE'), valeur(GetControlText('APL_QTEREALISE'))));
end;

procedure TOM_AFPLANNING.CreateControls;
begin
  fCB_APL_UNITETEMPS  := THDBValComboBox(GetControl('APL_UNITETEMPS'));
  fED_ATA_NUMEROTACHE := THEdit(GetControl('ATA_NUMEROTACHE'));
  fBT_Ferme           := TToolBarButton97(GetControl('bFerme'));
  fBT_Valider         := TToolBarButton97(GetControl('bValider'));
  fFO_PLANNING        := TForm(GetControl('AF_PLANNING'));
  fED_RESSOURCE       := THDBEdit(GetControl('APL_RESSOURCE'));

  fED_DATEDEBPLA      := THDBEdit(GetControl('APL_DATEDEBPLA'));
  fED_DATEFINPLA      := THDBEdit(GetControl('APL_DATEFINPLA'));
  fED_QTEPLANIFIEE    := THDBEdit(GetControl('APL_QTEPLANIFIEE'));

  fED_QTEPLANIFIEE    := THDBEdit(GetControl('APL_QTEPLANIFIEE'));
  fED_QTEREALISE      := THDBEdit(GetControl('APL_QTEREALISE'));

  fED_DATEDEBREAL      := THDBEdit(GetControl('APL_DATEDEBREAL'));
  fED_DATEFINREAL      := THDBEdit(GetControl('APL_DATEFINREAL'));
  fED_QTEREALISE       := THDBEdit(GetControl('APL_QTEREALISE'));

  // tant que le planning heure n'est pas géré
  SetControlVisible('GB_DATE', true);
  SetControlVisible('GB_HEURE', false);
  SetControlVisible('APL_ACTIVITEEFFECT', false);
                                         
end;

procedure TOM_AFPLANNING.LoadRecord;
begin

  // initialisation
  TPageControl(GetControl('Pages')).ActivePage := TTabSheet(GetControl('PGeneral'));

  // branchement apres le load
  fED_DATEDEBPLA.onExit := DATEDEBPLAExit;
  fED_DATEFINPLA.onExit := DATEFINPLAExit;
  fED_DATEDEBREAL.onExit := DATEDEBREALExit;
  fED_DATEFINREAL.onExit := DATEFINREALExit;
  fED_QTEPLANIFIEE.onExit := QTEPLANIFIEEExit;
  fED_QTEREALISE.onExit := QTEREALISEExit;
  fED_QTEPLANIFIEE.OnChange := APL_QTEPLANIFIEEOnChange;
  fED_QTEREALISE.OnChange := APL_QTEREALISEOnChange;
  fCB_APL_UNITETEMPS.OnChange := APL_UNITETEMPSOnChange;
  fED_RESSOURCE.OnChange := APL_RESSOURCEOnChange;
             
end;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 14/10/2002
Modifié le ... :   /  /
Description .. : filtre des taches
Mots clefs ... :
*****************************************************************}
Function TOM_AFPLANNING.PlusAffaire : string;
var
  vQr   : TQuery;
  vSt   : String;
  vNb   : Integer;
  i     : Integer;
  vTob  : Tob;

begin
  result := ' AND ATA_AFFAIRE = "' + fStAffaire + '"';
  if fStFonction <> '' then
    result := result + ' AND ATA_FONCTION = "' + fStFonction + '"';

  // recherche de l'ensemble des taches auquelles est affectée la ressource
  vSt := 'SELECT ATR_NUMEROTACHE FROM TACHERESSOURCE WHERE ATR_AFFAIRE = "' + fStAffaire + '"';
  vSt := vSt + ' AND ATR_RESSOURCE ="' + fED_RESSOURCE.Text + '"';
  vQr := nil;
  vTob := TOB.Create('TACHERESSOURCE', nil, -1);
  try
    vQR := OpenSql(vSt, True);
    vNb := 0;
    if Not vQR.Eof then
      begin

        vTob.LoadDetailDB('TACHERESSOURCE','','', vQR, False, True);
        for i := 0 to vTob.detail.count -1 do
          begin
            if vNb = 0 then result := result + ' AND (ATA_NUMEROTACHE = ' + floatToStr(vTob.detail[i].GetValue('ATR_NUMEROTACHE'))
            else result := result + ' OR ATA_NUMEROTACHE = ' + floatToStr(vTob.detail[i].GetValue('ATR_NUMEROTACHE'));
            vNb := vNb + 1;
          end;
      end;
    if vNb <> 0 then result := result + ')';
    fStPlusAffaire := result;
  finally
    ferme(vQr);
    vTob.Free;
  end;
end;

procedure TOM_AFPLANNING.UpdateMontant;
begin
  SetField('APL_DEVISE', fMaTob.getValue('APL_DEVISE'));
  SetField('APL_PUPR', fMaTob.getValue('APL_PUPR'));
  SetField('APL_PUPRDEV', fMaTob.getValue('APL_PUPRDEV'));
  SetField('APL_PUVENTEHT', fMaTob.getValue('APL_PUVENTEHT'));
  SetField('APL_PUVENTEDEVHT', fMaTob.getValue('APL_PUVENTEDEVHT'));
  SetField('APL_INITPTPR', fMaTob.getValue('APL_INITPTPR'));
  SetField('APL_INITPTPRDEV', fMaTob.getValue('APL_INITPTPRDEV'));
  SetField('APL_INITPTVENTEHT', fMaTob.getValue('APL_INITPTVENTEHT'));
  SetField('APL_INITPTVTDEVHT', fMaTob.getValue('APL_INITPTVTDEVHT'));
  SetField('APL_REALPTPR', fMaTob.getValue('APL_REALPTPR'));
  SetField('APL_REALPTPRDEV', fMaTob.getValue('APL_REALPTPRDEV'));
  SetField('APL_REALPTVENTEHT', fMaTob.getValue('APL_REALPTVENTEHT'));
  SetField('APL_REALPTVTDEVHT', fMaTob.getValue('APL_REALPTVTDEVHT'));
end;

procedure TOM_AFPLANNING.bMemoOnClick(SEnder: TObject);
begin
  AglLanceFiche('YY','YYLIENSOLE','APL;'+ GetControlText('APL_AFFAIRE') + '/' + GetControlText('APL_NUMEROLIGNE'),'', 'ACTION=' + fStAction);
end;

procedure TOM_AFPLANNING.SetArguments(StSQL: string);
var Critere,ChampMul,ValMul : string ;
    x,y : integer ;
    Ctrl : TControl;
    Fiche : TFFiche;
begin
SetControlVisible('BSTOP',TRUE);
DS.Edit;
Fiche := TFFiche(ecran);
Repeat
    Critere:=AnsiUppercase(Trim(ReadTokenPipe(StSQL,'|'))) ;
    if Critere<>'' then
        begin
        x:=pos('=',Critere);
        if x<>0 then
           begin
           ChampMul:=copy(Critere,1,x-1);
           ValMul:=copy(Critere,x+1,length(Critere));
           y := pos(',',ValMul);
           if y<>0 then ValMul:=copy(ValMul,1,length(ValMul)-1);
           if copy(ValMul,1,1)='"' then ValMul:=copy(ValMul,2,length(ValMul));
           if copy(ValMul,length(ValMul),1)='"' then ValMul:=copy(ValMul,1,length(ValMul)-1);
           SetField(ChampMul,ValMul);
           Ctrl:=TControl(Fiche.FindComponent(ChampMul));
           if Ctrl=nil then exit;
{$IFDEF EAGLCLIENT}
           if (Ctrl is TCustomCheckBox) or (Ctrl is THValComboBox) Or (Ctrl is TCustomEdit) then TEdit(Ctrl).Font.Color:=clRed
           else if Ctrl is TSpinEdit then TSpinEdit(Ctrl).Font.Color:=clRed
           else if (Ctrl is TCheckBox) or (Ctrl is THValComboBox) Or (Ctrl is THEdit)Or (Ctrl is THNumEdit)then
              begin
              TSpinEdit(Ctrl).Font.Color:=clRed;
              SetControlText(ChampMul,ValMul);
              end;
{$ELSE}
           if (Ctrl is TDBCheckBox) or (Ctrl is THDBValComboBox) Or (Ctrl is THDBEdit) then TEdit(Ctrl).Font.Color:=clRed
           else if Ctrl is THDBSpinEdit then THDBSpinEdit(Ctrl).Font.Color:=clRed
           else if (Ctrl is TCheckBox) or (Ctrl is THValComboBox) Or (Ctrl is THEdit)Or (Ctrl is THNumEdit)then
              begin
              THDBSpinEdit(Ctrl).Font.Color:=clRed;
              SetControlText(ChampMul,ValMul);
              end;
{$ENDIF}
           end;
        end;
until  Critere='';
end;

Initialization
  registerclasses ( [ TOM_AFPLANNING ] ) ;
end.

