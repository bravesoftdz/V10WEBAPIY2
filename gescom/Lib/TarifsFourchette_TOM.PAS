{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 13/06/2002
Modifié le ... :   /  /
Description .. : Source TOM de la TABLE : TarifsFourchette (TarifsFourchette)
Mots clefs ... : TOM;TarifsFourchette
*****************************************************************}
Unit TarifsFourchette_TOM ;

Interface

Uses
  StdCtrls,
  Controls,
  Classes,
  {$IFNDEF EAGLCLIENT}
    db,
    dbTables,
    Fe_Main,
    Fiche,
    FichList,
  {$ELSE}
    MainEagl,
    eFiche,
    eFichList,
  {$ENDIF}
  forms,
  sysutils,
  ComCtrls,
  Graphics,
  HCtrls,
  HEnt1,
  HMsgBox,
  UTOM,
  UTob,
  wTom,
  SaisieList,
  uTableFiltre;

Type
  TOM_yTarifsFourchette = Class (twTOM)
  procedure OnNewRecord                ; override ;
  procedure OnDeleteRecord             ; override ;
  procedure OnUpdateRecord             ; override ;
  procedure OnAfterUpdateRecord        ; override ;
  procedure OnLoadRecord               ; override ;
  procedure OnChangeField ( F: TField) ; override ;
  procedure OnArgument ( S: String )   ; override ;
  procedure OnClose                    ; override ;
  procedure OnCancelRecord             ; override ;
  private
    iIdentifiantYTA  : integer;
    iIdentifiantYTQ  : integer;
    sFonctionnalite  : string;
    sMethodeDeSaisie : string;
    sAction, sDroit  : string;
    sDevise          : string;
    sNatureAuxi      : string;
    nMini, nMaxi     : double;
    lAfterDelete     : boolean;
    iSensTaux        : integer;
    TF : TTAbleFiltre;
    procedure SetSayOuGet(sControl:String; lConditionSay:boolean);
    function  ControleAffichagePrixBrutPrixNet : boolean;
    procedure ControleCoherenceFourchette(sOu:string);

    function wWhereYTF(iIdentifiantYTA, iIdentifiantYTQ : integer): string;
    function wExistYTF(iIdentifiantYTA, iIdentifiantYTQ : integer):Boolean;

  protected
    { Control Field }
    procedure ControlField(FieldName: string);
    procedure CalculField(FieldName: string);
    Function  RecordIsValid: boolean; override;
  end ;

Implementation

uses
  wCommuns
  ,Tarifs
  ,wJetons
  ;

Const
  { Tableau des libellés de messages pour les ressources }
  TexteMessage: array[1..2] of string =
      (
      {1}  'Les fourchettes sont mal définies (De >= A)',
      { 2}  'Le taux de remise ne doit pas dépasser 100%'
      );

  ErrFourchette  = 1;
  ErrRemise100   = 2;
  n0             =          0.00;
  n9999999999    = 9999999999.00;

{--------------------------------------------------------------------------------
---------------------------------------------------------------------------------}
procedure TOM_yTarifsFourchette.OnNewRecord ;
begin
  Inherited ;
  if (Ecran<>nil) then
  begin
    SetInteger('YTF_IDENTIFIANTYTA',iIdentifiantYTA);
    SetInteger('YTF_IDENTIFIANTYTQ',iIdentifiantYTQ);
    // Si aucune ligne
    if (not wExistYTF(iIdentifiantYTA, iIdentifiantYTQ) ) then
    begin
      TF.Insert;
      TF.PutValue('YTF_IDENTIFIANTYTA',iIdentifiantYTA);
      TF.PutValue('YTF_IDENTIFIANTYTQ',iIdentifiantYTQ);
      TF.PutValue('YTF_FOURCHETTEMINI',n0);
      TF.PutValue('YTF_FOURCHETTEMAXI',n9999999999);
      TF.Post;
    end
    else
    begin
      SetDouble('YTF_FOURCHETTEMINI',iif(nMaxi=n9999999999,nMini,nMaxi));
    end;
  end;
end ;

{--------------------------------------------------------------------------------
---------------------------------------------------------------------------------}
procedure TOM_yTarifsFourchette.OnDeleteRecord ;
begin
  Inherited ;
  lAfterDelete := True;
end ;

{--------------------------------------------------------------------------------
---------------------------------------------------------------------------------}
procedure TOM_yTarifsFourchette.OnUpdateRecord ;
begin
  DisableControl;
  try

  finally
    EnableControl;
  end;

  if RecordIsValid then
  begin

  end;

  Inherited ;
end ;

{--------------------------------------------------------------------------------
---------------------------------------------------------------------------------}
procedure TOM_yTarifsFourchette.OnAfterUpdateRecord ;
begin
  Inherited ;
  ControleCoherenceFourchette('Après Saisie');
end ;

{--------------------------------------------------------------------------------
---------------------------------------------------------------------------------}
procedure TOM_yTarifsFourchette.OnLoadRecord ;
begin
  Inherited ;
  nMini:= GetDouble('YTF_FOURCHETTEMINI');
  nMaxi:= GetDouble('YTF_FOURCHETTEMAXI');
  ControleAffichagePrixBrutPrixNet;
  if (lAfterDelete) then
  begin
    ControleCoherenceFourchette('Après Annulation');
  end;
end;

{--------------------------------------------------------------------------------
---------------------------------------------------------------------------------}
procedure TOM_yTarifsFourchette.OnChangeField ( F: TField ) ;
begin
  Inherited ;
end ;

{--------------------------------------------------------------------------------
---------------------------------------------------------------------------------}
procedure TOM_yTarifsFourchette.OnArgument ( S: String ) ;
var
  lSaisieEnGrille : boolean;
begin
  Inherited ;
  iIdentifiantYTA   := ValeurI(GetArgumentValue(S, 'YTF_IDENTIFIANTYTA'));
  iIdentifiantYTQ   := ValeurI(GetArgumentValue(S, 'YTF_IDENTIFIANTYTQ'));
  sDevise           := GetArgumentValue(S,'YTA_DEVISE');
  sFonctionnalite   := GetArgumentValue(S,'YTA_FONCTIONNALITE');
  sNatureAuxi       := NatureAuxiliaire(sFonctionnalite);
  sMethodeDeSaisie  := GetArgumentValue(S,'METHODEDESAISIE');
  sAction           := GetArgumentValue(S,'ACTION');
  sDroit            := GetArgumentValue(S,'DROIT');
  lAfterDelete      := False;

  Ecran.Caption := RechDom('YFONCTIONNALITES',sFonctionnalite,False)+TraduireMemoire(' : Fourchettes');

  if (Ecran<>nil) then
  begin
    TF := TFSaisieList(Ecran).LeFiltre;
    lSaisieEnGrille := (sMethodeDeSaisie='GRILLE');
    TF.SaisieMixte  := (sMethodeDeSaisie='MIXTE');
    TF.SaisieEnGrid := ((sMethodeDeSaisie='MIXTE') or (sMethodeDeSaisie='GRILLE'));
    TFSaisieList(Ecran).PCPied.Visible        := not lSaisieEnGrille;
    TFSaisieList(Ecran).PanPied.visible       := not lSaisieEnGrille;
    TFSaisieList(Ecran).SplitMessages.Visible := not lSaisieEnGrille;
  end;

  //Les taux sont-ils à retrancher ( remise ) ou à augmenter ( taux commissionnement par exemple )
  if (sFonctionnalite=sCommissionclient) then
    iSensTaux := +1
  else
    iSensTaux := -1;

  //Fixer les libellés par rapport à la fonctionnalité
  if (GetControl('TYTF_FORFAIT')       <>nil) then tHLabel(GetControl('TYTF_FORFAIT')      ).Caption := LibelleParFonctionnalite(sFonctionnalite, 'FORFAIT'      );
  if (GetControl('TYTF_COUTLANCEMENT') <>nil) then tHLabel(GetControl('TYTF_COUTLANCEMENT')).Caption := LibelleParFonctionnalite(sFonctionnalite, 'COUTLANCEMENT');
  if (GetControl('TYTF_REMISEMONTANT') <>nil) then tHLabel(GetControl('TYTF_REMISEMONTANT')).Caption := LibelleParFonctionnalite(sFonctionnalite, 'REMISEMONTANT');
  if (GetControl('TYTF_PRIXBRUT')      <>nil) then tHLabel(GetControl('TYTF_PRIXBRUT')     ).Caption := LibelleParFonctionnalite(sFonctionnalite, 'PRIXBRUT'     );
  if (GetControl('TYTF_REMISE1')       <>nil) then tHLabel(GetControl('TYTF_REMISE1')      ).Caption := LibelleParFonctionnalite(sFonctionnalite, 'REMISE123'    );
  if (GetControl('TYTF_PRIXNET')       <>nil) then tHLabel(GetControl('TYTF_PRIXNET')      ).Caption := LibelleParFonctionnalite(sFonctionnalite, 'PRIXNET'      );
  if (GetControl('TYTF_CONDFORFAIT')   <>nil) then tHLabel(GetControl('TYTF_CONDFORFAIT')  ).Caption := LibelleParFonctionnalite(sFonctionnalite, 'FORFAIT'      );
  if (GetControl('TYTF_CONDCOUTLANC')  <>nil) then tHLabel(GetControl('TYTF_CONDCOUTLANC') ).Caption := LibelleParFonctionnalite(sFonctionnalite, 'COUTLANCEMENT');
  if (GetControl('TYTF_CONDREMISEM')   <>nil) then tHLabel(GetControl('TYTF_CONDREMISEM')  ).Caption := LibelleParFonctionnalite(sFonctionnalite, 'REMISEMONTANT');
  if (GetControl('TYTF_CONDPRIXBRUT')  <>nil) then tHLabel(GetControl('TYTF_CONDPRIXBRUT') ).Caption := LibelleParFonctionnalite(sFonctionnalite, 'PRIXBRUT'     );
  if (GetControl('TYTF_CONDREMISE1')   <>nil) then tHLabel(GetControl('TYTF_CONDREMISE1')  ).Caption := LibelleParFonctionnalite(sFonctionnalite, 'REMISE123'    );
  if (GetControl('TYTF_CONDPRIXNET')   <>nil) then tHLabel(GetControl('TYTF_CONDPRIXNET')  ).Caption := LibelleParFonctionnalite(sFonctionnalite, 'PRIXNET'      );

end ;

{--------------------------------------------------------------------------------
---------------------------------------------------------------------------------}
procedure TOM_yTarifsFourchette.OnClose ;
begin
  Inherited ;
end ;

{--------------------------------------------------------------------------------
---------------------------------------------------------------------------------}
procedure TOM_yTarifsFourchette.OnCancelRecord ;
begin
  Inherited ;
end ;

{--------------------------------------------------------------------------------*
   Data Change Champ par Champ
---------------------------------------------------------------------------------}
procedure TOM_yTarifsFourchette.CalculField(FieldName: string);
begin
  if (FieldName='YTF_FOURCHETTEMINI') then
    SetSayOuGet('YTF_FOURCHETTEMINI' ,GetDouble('YTF_FOURCHETTEMINI')=n0         )
  else if (FieldName='YTF_FOURCHETTEMAXI') then
    SetSayOuGet('YTF_FOURCHETTEMAXI' ,GetDouble('YTF_FOURCHETTEMAXI')=n9999999999);
end;

{--------------------------------------------------------------------------------*
   Validation Champ par Champ
---------------------------------------------------------------------------------}
procedure TOM_yTarifsFourchette.ControlField(FieldName: string);
begin
	DisableControl;
  try
    if      (FieldName = 'YTF_FORFAIT')         then SetBoolean('YTF_CONDFORFAIT' ,GetDouble('YTF_FORFAIT')       <>0.0)
    else if (FieldName = 'YTF_COUTLANCEMENT')   then SetBoolean('YTF_CONDCOUTLANC',GetDouble('YTF_COUTLANCEMENT') <>0.0)
    else if (FieldName = 'YTF_REMISEMONTANT')   then SetBoolean('YTF_CONDREMISEM' ,GetDouble('YTF_REMISEMONTANT') <>0.0)
    else if (FieldName = 'YTF_PRIXBRUT')        then SetBoolean('YTF_CONDPRIXBRUT',GetDouble('YTF_PRIXBRUT')      <>0.0)
    else if (FieldName = 'YTF_REMISE1')         then SetBoolean('YTF_CONDREMISE1' ,GetDouble('YTF_REMISE1')       <>0.0)
    else if (FieldName = 'YTF_REMISE2')         then SetBoolean('YTF_CONDREMISE2' ,GetDouble('YTF_REMISE2')       <>0.0)
    else if (FieldName = 'YTF_REMISE3')         then SetBoolean('YTF_CONDREMISE3' ,GetDouble('YTF_REMISE3')       <>0.0)
    else if (FieldName = 'YTF_PRIXNET')         then SetBoolean('YTF_CONDPRIXNET' ,GetDouble('YTF_PRIXNET')       <>0.0)
    ;

    if (Pos(FieldName,'YTF_PRIXBRUT/YTF_REMISE1/YTF_REMISE2/YTF_REMISE3/YTF_PRIXNET')>0) then
    begin
      if (Pos(FieldName,'YTF_REMISE1')>0) and (GetDouble('YTF_REMISE1')>100) then LastError := ErrRemise100;
      if (Pos(FieldName,'YTF_REMISE2')>0) and (GetDouble('YTF_REMISE2')>100) then LastError := ErrRemise100;
      if (Pos(FieldName,'YTF_REMISE3')>0) and (GetDouble('YTF_REMISE3')>100) then LastError := ErrRemise100;

      if (ControleAffichagePrixBrutPrixNet) then
        SetDouble('YTF_PRIXNET',CalculPrixNet( iif(GetBoolean('YTF_CONDPRIXBRUT'),GetDouble('YTF_PRIXBRUT'),0)
                                              ,iif(GetBoolean('YTF_CONDREMISE1') ,GetDouble('YTF_REMISE1') ,0)
                                              ,iif(GetBoolean('YTF_CONDREMISE2') ,GetDouble('YTF_REMISE2') ,0)
                                              ,iif(GetBoolean('YTF_CONDREMISE3') ,GetDouble('YTF_REMISE3') ,0)
                                              ,iSensTaux,sDevise,sNatureAuxi,'P','P'));
    end;
  finally
    EnableControl;
  end;

//  inherited;

  if LastError <> 0 then
    LastErrorMsg := TexteMessage[LastError];
end;

{--------------------------------------------------------------------------------*
   Validation Global
---------------------------------------------------------------------------------}
function TOM_yTarifsFourchette.RecordIsValid: boolean;
begin
  if      (GetDouble('YTF_FOURCHETTEMINI')>=GetDouble('YTF_FOURCHETTEMAXI')) then LastError := ErrFourchette
  ;
  Result := LastError = 0;

  if (LastError>0) then
  begin
  LastErrorMsg := TexteMessage[LastError];
    Case LastError of
      ErrFourchette : SetFocusControl('YTF_FOUCHETTEMAXI');
    end;
  end;
end;

{--------------------------------------------------------------------------------*
   Controle et Affichage des zone Prix Brute / Remise 1/2/3 et Prix Net
---------------------------------------------------------------------------------}
function TOM_yTarifsFourchette.ControleAffichagePrixBrutPrixNet : boolean;
var
  lPrixBrut, lNiLunNiLautre : boolean;
begin
  lPrixBrut      := (
                        ((GetDouble('YTF_PRIXBRUT')<>0) and (GetBoolean('YTF_CONDPRIXBRUT')))
                     or ((GetDouble('YTF_REMISE1') <>0) and (GetBoolean('YTF_CONDREMISE1') ))
                     or ((GetDouble('YTF_REMISE2') <>0) and (GetBoolean('YTF_CONDREMISE2') ))
                     or ((GetDouble('YTF_REMISE3') <>0) and (GetBoolean('YTF_CONDREMISE3') ))
                    );
  lNiLunNiLautre := (not lPrixBrut) and (GetDouble('YTF_PRIXNET')=0);
  SetSayOuGet('YTF_FORFAIT'        , False);
  SetSayOuGet('YTF_COUTLANCEMENT'  , False);
  SetSayOuGet('YTF_REMISEMONTANT'  , False);
  SetSayOuGet('YTF_PRIXNET'        , (    lPrixBrut) and (not lNiLunNiLautre));
  SetSayOuGet('YTF_PRIXBRUT'       , (not lPrixBrut) and (not lNiLunNiLautre));
  SetSayOuGet('YTF_REMISE1'        , (not lPrixBrut) and (not lNiLunNiLautre));
  SetSayOuGet('YTF_REMISE2'        , (not lPrixBrut) and (not lNiLunNiLautre));
  SetSayOuGet('YTF_REMISE3'        , (not lPrixBrut) and (not lNiLunNiLautre));
  Result := lPrixBrut or lNiLunNiLautre;
end;

{-------------------------------------------------------------------------------
   Positionne les propriété Enable et Color d'un controle en fonction d'une condition de Say
--------------------------------------------------------------------------------}
procedure TOM_yTarifsFourchette.SetSayOuGet(sControl:String; lConditionSay:boolean);
begin
  if lConditionSay or (sAction='CONSULTATION') then
  begin
    SetControlProperty(sControl, 'Enabled', False);
    SetControlProperty(sControl, 'Color'  , ClBtnFace);
  end
  else
  begin
    SetControlProperty(sControl, 'Enabled', True);
    SetControlProperty(sControl, 'Color'  , ClWindow);
  end
end;

{--------------------------------------------------------------------------------*
   Teste existance des lignes de tarifs TarifsFourchette
---------------------------------------------------------------------------------}
function TOM_yTarifsFourchette.wExistYTF(iIdentifiantYTA, iIdentifiantYTQ : integer):Boolean;
var
	sSql  : string;
begin
	sSql := ' SELECT YTF_IDENTIFIANT'
			+ ' FROM yTarifsFourchette'
			+ ' WHERE ' + wWhereYTF(iIdentifiantYTA, iIdentifiantYTQ)
         ;
	Result := ExisteSQL(sSql);
end;

{--------------------------------------------------------------------------------*
   Construction du WHERE pour la fiche TIERS
---------------------------------------------------------------------------------}
function TOM_yTarifsFourchette.wWhereYTF(iIdentifiantYTA, iIdentifiantYTQ : integer): string;
begin
	Result := 'YTF_IDENTIFIANTYTA = "' + IntToStr(iIdentifiantYTA) + '" and '+'YTF_IDENTIFIANTYTQ = "' + IntToStr(iIdentifiantYTQ) + '"';
end;

{--------------------------------------------------------------------------------
   Controle cohérence et Mise à Jour au cas où
---------------------------------------------------------------------------------}
procedure TOM_yTarifsFourchette.ControleCoherenceFourchette(sOu:string);
var
  TheTob : TOB;
begin
  lAfterDelete := False;
  if      (sOu='Après Annulation') then
  begin
    if (TF<>nil) then
    begin
      TF.StartUpdate;
      TF.DisableTOM;

      TF.first;
      nMaxi:=TF.GetValue('YTF_FOURCHETTEMAXI');

      if (TF.GetValue('YTF_FOURCHETTEMINI')<>n0) then
      begin
        TF.PutValue('YTF_FOURCHETTEMINI',n0);
        TF.Post;
      end;

      TF.Next;
      while (not TF.eof) do
      begin
        if (TF.GetValue('YTF_FOURCHETTEMINI')<>nMaxi) then
        begin
          TF.PutValue('YTF_FOURCHETTEMINI',nMaxi);
          TF.Post;
        end;
        nMaxi:=TF.GetValue('YTF_FOURCHETTEMAXI');
        TF.Next;
      end;

      if (TF.GetValue('YTF_FOURCHETTEMAXI')<>n9999999999) then
      begin
        TF.PutValue('YTF_FOURCHETTEMAXI',n9999999999);
        TF.Post;
      end;

      TF.EnableTOM;
      TF.EndUpdate;
    end;
  end
  else if (sOu='Après Saisie') then
  begin
    if (TF<>nil) then
    begin
      //Il faut retrier le tableau sur la clef primaire puis se repositionner sur la ligne en cours de modification
      TheTob := TF.TOBFiltre.Detail[TF.Recno-1];
      TF.RefreshControls;

      TF.StartUpdate; //Déconnecte les controles de l'écran
      TF.DisableTOM;  //Déconnecte les controles de la TOM

      //On recupère les données de la ligne modifiée
      nMini  := GetDouble('YTF_FOURCHETTEMINI');
      nMaxi  := GetDouble('YTF_FOURCHETTEMAXI');

      //Mise en cohérence des enregistrements précédents par rapport à celui saisi
      TF.prior;
      while (not TF.bof) do
      begin
        if (TF.GetValue('YTF_FOURCHETTEMINI')<nMini) then
        begin
          TF.PutValue('YTF_FOURCHETTEMAXI',nMini);
          TF.Post;
          break;
        end
        else
        begin
          TF.delete;
        end;
        TF.Prior;
      end;

      //Il faut retrier le tableau sur la clef primaire à cause de delete éventuel puis se repositionner sur la ligne en cours de modification
      TF.RefreshControls;
      TF.SelectRecord( TheTob.GetIndex+1 );

      //Mise en cohérence des enregistrements suivants par rapport à celui saisi
      TF.Next;
      while (not TF.eof) do
      begin
        if (TF.GetValue('YTF_FOURCHETTEMAXI')>nMaxi) then
        begin
          TF.PutValue('YTF_FOURCHETTEMINI',nMaxi);
          TF.Post;
          break;
        end
        else
        begin
          TF.delete;
        end;
        TF.Next;
      end;
      TF.EnableTOM;   //Reconnecte les controles de la TOM
      TF.EndUpdate;   //Reconnecte les controles de l'écran
    end;
  end;
end;

Initialization
  registerclasses ( [ TOM_yTarifsFourchette ] ) ;
end.
