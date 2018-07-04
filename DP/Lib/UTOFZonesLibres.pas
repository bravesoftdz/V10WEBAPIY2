{***********UNITE*************************************************
Auteur  ...... : BM
Créé le ...... : 05/11/2002
Modifié le ... :   /  /
Description .. : Source TOF_ZONESLIBRES : initialisation des onglets "Zones libres"
Mots clefs ... : TOF;ZONESLIBRES
*****************************************************************}

Unit UTOFZonesLibres;

Interface

Uses
   {$IFNDEF CJS5}
   utofaftraducchamplibre,
   {$ENDIF}
   HEnt1,
   UTOF, UTOB, hctrls;

////////////////////////////// CLASSE //////////////////////////////
Type
   {$IFDEF CJS5}
   TOF_ZONESLIBRES = Class (TOF)
   {$ELSE}
   TOF_ZONESLIBRES = Class (TOF_AFTRADUCCHAMPLIBRE)
   {$ENDIF}
      private
         sNomPanel_f : string;
         sTableParam_f : string;
         sNomTablette_f : string;
         bOnglet_f : boolean;
         nOngletNb_f : integer;

         procedure ChargeCaptionsParam;
         function  TraiteLabel (hLabel_p:THLabel; OBTablette_p,OBCaption_p:TOB):integer;

      public
         procedure OnLoad; override;
         procedure OnArgument (StrArguments : String) ; Override;

         procedure ZonesLibresOnglets (sNomPanel_p:string);
         procedure ZonesLibresInit    (sTableParam_p,sNomTablette_p:string);
   end;


////////////////////////////// IMPLEMENTATION //////////////////////////////
Implementation

uses
   {$IFNDEF EAGLCLIENT}
   Mul, 
   {$ELSE}
   eMul,
   {$ENDIF}
   Classes, SysUtils, hmsgbox, controls, comctrls, ParamSoc;


{*****************************************************************
Auteur ....... : B. MERIAUX
Date ......... : 12/12/02
Procédure .... : OnLoad
Description .. : Vérifie l'initialisation de la table et de la
                 tablette de paramétrage et charge les libellé
                 des labels
Paramètres ... :
*****************************************************************}
procedure TOF_ZONESLIBRES.OnLoad ;
begin
   Inherited ;

     
   if not bOnglet_f then
      exit;

   if  ( sNomPanel_f = '' ) then
   begin
      PGIInfo( 'Nom du panel non défini.#13#10' +
               'Utilisez la fonction "ZonesLibresOnglets" dans le "OnArgument" de votre TOF',
               'Initialisation des onglets "Zones Libres"' );
      exit;
   end;

   if  ( sTableParam_f = '' ) or ( sNomTablette_f = '' ) then
   begin
      PGIInfo( 'Table ou tablette de paramétrage non définies#13#10' +
               'Utilisez la fonction "ZonesLibresInit" dans le "OnArgument" de votre TOF',
               'Initialisation des onglets "Zones Libres"' );
      exit;
   end;

   ChargeCaptionsParam;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : B.MERIAUX
Créé le ...... : 02/09/2003
Modifié le ... :   /  /
Description .. : Initialise le nom des onglets et les affiche
Mots clefs ... : 
*****************************************************************}
procedure TOF_ZONESLIBRES.ZonesLibresOnglets( sNomPanel_p : string);
begin
   sNomPanel_f := sNomPanel_p;
   nOngletNb_f := 0;
   While ( GetControl(sNomPanel_f + IntToStr(nOngletNb_f + 1) ) <> nil) do
   begin
      SetControlVisible( sNomPanel_f + IntToStr(nOngletNb_f + 1), false );
      inc(nOngletNb_f);
   end;

   bOnglet_f := (nOngletNb_f <> 0);
end;

{*****************************************************************
Auteur ....... : B. MERIAUX
Date ......... : 12/12/02
Procédure .... : ZonesLibresInit
Description .. : Initialise la table et la tablette de paramétrage
                 si un ou plusieurs onglets existent
Paramètres ... : La table, la tablette
*****************************************************************}
procedure TOF_ZONESLIBRES.ZonesLibresInit( sTableParam_p, sNomTablette_p : string);
begin
   if not bOnglet_f then
      exit;
   sTableParam_f := sTableParam_p;
   sNomTablette_f := sNomTablette_p;
end;

{*****************************************************************
Auteur ....... : B. MERIAUX
Date ......... : 12/12/02
Procédure .... : IncCompteur
Description .. : Charge les libellés des labels
Paramètres ... :
*****************************************************************}
procedure TOF_ZONESLIBRES.ChargeCaptionsParam;
var
   OBCaption_l, OBTablette_l : TOB;
   nFICompNb_l, nOngletInd_l : integer;
   sOngletNom_l, sTable_l : string;
   tnCompOngletNb_l : array of integer;
begin
   // Pour les tablettes
   sTable_l := PrefixeToTable(TTToPrefixe( sNomTablette_f));
   OBTablette_l := TOB.Create( sNomTablette_f, nil, -1);
   OBTablette_l.LoadDetailDB( sTable_l, '"' + TTToTipe( sNomTablette_f) + '"',
                              TTToCode( sNomTablette_f), nil, false );

   // Pour les autres champs
   OBCaption_l := TOB.Create(sTableParam_f, nil, -1);
   OBCaption_l.SelectDB('',nil);

   SetLength(tnCompOngletNb_l, nOngletNb_f);

   For nFICompNb_l := 0 to TFMul(Ecran).ComponentCount -1 do
   begin
      if ( Ecran.Components[nFICompNb_l] is TControl ) and
         ( Ecran.Components[nFICompNb_l].ClassName = 'THLabel' ) and
         ( TControl(Ecran.Components[nFICompNb_l]).Parent is TTabSheet ) then
      begin
         sOngletNom_l := TTabSheet( TControl(Ecran.Components[nFICompNb_l]).Parent ).Name;
         if Copy(sOngletNom_l, 1, Length(sOngletNom_l) - 1) = sNomPanel_f then
         begin
            nOngletInd_l := StrToInt(sOngletNom_l[Length(sOngletNom_l)]) - 1;
            tnCompOngletNb_l[nOngletInd_l] := tnCompOngletNb_l[nOngletInd_l] +
                  TraiteLabel( THLabel(Ecran.Components[nFICompNb_l]), OBTablette_l, OBCaption_l );
         end;
      end;
   end;

   for nOngletInd_l := 1 to nOngletNb_f do
      SetControlVisible( sNomPanel_f + IntToStr(nOngletInd_l) , (tnCompOngletNb_l[nOngletInd_l-1] <> 0) );

   OBCaption_l.Free;
   OBTablette_l.Free;
end ;

{*****************************************************************
Auteur ....... : B. MERIAUX
Date ......... : 12/12/02
Fonction ..... : TraiteLabel
Description .. : Détermine le libellé du label, puis si le label
                 et son champ associé doivent être affichés
Paramètres ... : Le label, la TOB tablette, la TOB caption
Renvoie ...... : 0 si le préfixe du label ne correspond pas à celui de la table de paramétrage
                 0 si le label ou le champ associé n'existent pas
                 0 si le label ou le champ associé ne sont pas affichés
                 1 sinon
*****************************************************************}
function TOF_ZONESLIBRES.TraiteLabel( hLabel_p: THLabel; OBTablette_p, OBCaption_p : TOB ) : integer;
var
   wcFocus_l : TWinControl;
   sPrefixe_l, sTablette_l, sValLabel_l, sNomLabel_l : string;
   nTabInd_l : integer;
begin
   result := 0;

   sPrefixe_l := Copy( hLabel_p.Name, 1, Pos( '_', hLabel_p.Name ) - 1 );
   if sPrefixe_l <> TableToPrefixe( sTableParam_f ) then
   begin
      SetControlCaption( hLabel_p.Name, '*** Pas défini ***' );
      exit;
   end;

   wcFocus_l := hLabel_p.FocusControl;
   if wcFocus_l = nil then
   begin
      SetControlCaption( hLabel_p.Name, '*** Pas de focus ***' );
      exit;
   end;

   if ( wcFocus_l.ClassName = 'THValComboBox' ) then
   begin
      sTablette_l := THValComboBox(wcFocus_l).DataType;
      nTabInd_l := StrToInt(sTablette_l[Length(sTablette_l)]) - 1;
      sValLabel_l := RechDom( sNomTablette_f, OBTablette_p.Detail[nTabInd_l].GetValue(  TTToCode( sNomTablette_f) ), false );
   end
   else if ( wcFocus_l.ClassName = 'THMultiValComboBox' ) then
   begin
      sTablette_l := THMultiValComboBox(wcFocus_l).DataType;
      nTabInd_l := StrToInt(sTablette_l[Length(sTablette_l)]) - 1;
      sValLabel_l := RechDom( sNomTablette_f, OBTablette_p.Detail[nTabInd_l].GetValue(  TTToCode( sNomTablette_f) ), false );
   end
   else
   begin
      sNomLabel_l := hLabel_p.Name;
      if sNomLabel_l[Length(sNomLabel_l)] = '_' then
         sNomLabel_l := Copy(sNomLabel_l, 1, Length(sNomLabel_l) - 1 );
      sValLabel_l := OBCaption_p.GetValue( sNomLabel_l );
   end;

   if hLabel_p.Name[Length(hLabel_p.Name)] <> '_' then
      SetControlCaption( hLabel_p.Name, sValLabel_l );

   SetControlVisible( hLabel_p.Name, (sValLabel_l <> '') );
   SetControlVisible( wcFocus_l.Name, (sValLabel_l <> '') );

   if (sValLabel_l <> '') then
      result := 1;
end ;

////////////////////////////// INITIALIZATION //////////////////////////////

procedure TOF_ZONESLIBRES.OnArgument(StrArguments: String);
begin
  inherited;

     // MB : Pour KPMG, ne pas chercher automatiquement
  if (Ecran is tfMUL) and (GetParamSocSecur ('SO_AFCLIENT', 0) <>8 ) then
        TFMul(Ecran).AutoSearch :=  asMouette ;

// RP 04/04/07 - Ajout dans cet objet car nécessaire pour dpTofMulAnnDoss, dpTofMulSelDoss, dpTOFAnnuSel
// $$$ JP 12/05/06 - POUR INEXTENSO
{$IFDEF BUREAU}
  if (GetControl('YTC_TABLELIBRETIERS1')<>Nil) and (GetParamSocSecur ('SO_LIBTIERSDEFMUL', FALSE) = FALSE) then
  begin
       SetControlText ('YTC_TABLELIBRETIERS1', '');
       SetControlText ('YTC_TABLELIBRETIERS2', '');
       SetControlText ('YTC_TABLELIBRETIERS3', '');
       SetControlText ('YTC_TABLELIBRETIERS4', '');
       SetControlText ('YTC_TABLELIBRETIERS5', '');
       SetControlText ('YTC_TABLELIBRETIERS6', '');
       SetControlText ('YTC_TABLELIBRETIERS7', '');
       SetControlText ('YTC_TABLELIBRETIERS8', '');
       SetControlText ('YTC_TABLELIBRETIERS9', '');
       SetControlText ('YTC_TABLELIBRETIERSA', '');
  end;
{$ENDIF}

end;

Initialization
   registerclasses ( [ TOF_ZONESLIBRES ] ) ;
end.


