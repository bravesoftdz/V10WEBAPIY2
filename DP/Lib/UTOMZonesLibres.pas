{***********UNITE*************************************************
Auteur  ...... : BM
Créé le ...... : 05/11/2002
Modifié le ... :   /  /
Description .. : Source TOF_ZONESLIBRES : initialisation des onglets "Zones libres"
Mots clefs ... : TOF;ZONESLIBRES
*****************************************************************}

Unit UTOMZonesLibres;

Interface

Uses
     UTOM, UTOB, hctrls, DBCtrls;


////////////////////////////// CLASSE //////////////////////////////

Type
   TOM_ZONESLIBRES = Class (TOM)
      private
         sTableParam_c : string;
         sNomTablette_c : string;
         bOnglet_c : boolean;

         procedure ChargeCaptionsParam;
         function TraiteLabel( hLabel_p: THLabel; OBTablette_p, OBCaption_p : TOB ) : integer;
         function TraiteCheckBox( hCheckBox_p: TDBCheckBox; OBCaption_p : TOB ) : integer;
      public
         procedure OnLoadRecord; override ;
         procedure OnArgument ( sParams_p : String )   ; override ;

         procedure ZonesLibresInit( sTableParam_p, sNomTablette_p : string);
   end ;

////////////////////////////// IMPLEMENTATION //////////////////////////////

Implementation

uses
   {$IFNDEF EAGLCLIENT}
   Fiche,
   {$ELSE}
   eFiche,
   {$ENDIF}
hmsgbox, controls, comctrls, SysUtils, Classes;

{*****************************************************************
Auteur ....... : B. MERIAUX
Date ......... : 12/12/02
Procédure .... : OnArgument
Description .. : Détermine et initialise le nombre d'onglets
Paramètres ... : Liste des arguments
*****************************************************************}
procedure TOM_ZONESLIBRES.OnArgument ( sParams_p: String ) ;
begin
   Inherited ;
   if ( GetControl('PDIVERS' ) <> nil) then
   begin
      SetControlVisible( 'PDIVERS', false );
      bOnglet_c := true;
   end
   else
      bOnglet_c := false;
end;

{*****************************************************************
Auteur ....... : B. MERIAUX
Date ......... : 12/12/02
Procédure .... : OnLoad
Description .. : Vérifie l'initialisation de la table et de la
                 tablette de paramétrage et charge les libellé
                 des labels
Paramètres ... :
*****************************************************************}
procedure TOM_ZONESLIBRES.OnLoadRecord ;
begin
   Inherited ;
   if not bOnglet_c then
      exit;

   if   ( sTableParam_c = '' ) and ( sNomTablette_c = '' ) then
   begin
      bOnglet_c := false;
//      PGIInfo( 'Table et tablette de paramétrage des zones libres non définies#13#10',
//              'Initialisation des onglets "Zones Libres"' );
      exit;
   end;

   ChargeCaptionsParam;
end ;

{*****************************************************************
Auteur ....... : B. MERIAUX
Date ......... : 12/12/02
Procédure .... : ZonesLibresInit
Description .. : Initialise la table et la tablette de paramétrage
                 si un ou plusieurs onglets existent
Paramètres ... : La table, la tablette
*****************************************************************}
procedure TOM_ZONESLIBRES.ZonesLibresInit( sTableParam_p, sNomTablette_p : string);
begin
   if not bOnglet_c then
      exit;

   sTableParam_c := sTableParam_p;
   sNomTablette_c := sNomTablette_p;
end;

{*****************************************************************
Auteur ....... : B. MERIAUX
Date ......... : 12/12/02
Procédure .... : IncCompteur
Description .. : Charge les libellés des labels
Paramètres ... :
*****************************************************************}
procedure TOM_ZONESLIBRES.ChargeCaptionsParam;
var
   OBCaption_l, OBTablette_l : TOB;
   nFICompNb_l, nCompOngletNb_l : integer;
   sOngletNom_l, sClasseNom_l, sTable_l : string;
begin
   // Pour les tablettes
   sTable_l := PrefixeToTable(TTToPrefixe( sNomTablette_c));
   OBTablette_l := TOB.Create( sNomTablette_c, nil, -1);
   OBTablette_l.LoadDetailDB( sTable_l, '"' + TTToTipe( sNomTablette_c) + '"',
                              TTToCode( sNomTablette_c), nil, false );

   // Pour les autres champs
   OBCaption_l := TOB.Create(sTableParam_c, nil, -1);
   OBCaption_l.SelectDB('',nil);

   nCompOngletNb_l := 0;

   For nFICompNb_l := 0 to TFFiche(Ecran).ComponentCount -1 do
   begin
      sClasseNom_l := Ecran.Components[nFICompNb_l].ClassName;
      if ( Ecran.Components[nFICompNb_l] is TControl ) and
         ( ( sClasseNom_l = 'THLabel' ) or ( sClasseNom_l = 'TDBCheckBox' ) ) and
         ( TControl(Ecran.Components[nFICompNb_l]).Parent is TTabSheet ) then
      begin
         sOngletNom_l := TTabSheet( TControl(Ecran.Components[nFICompNb_l]).Parent ).Name;
         if sOngletNom_l = 'PDIVERS' then
         begin
            if sClasseNom_l = 'THLabel' then
            begin
               nCompOngletNb_l := nCompOngletNb_l +
                     TraiteLabel( THLabel(Ecran.Components[nFICompNb_l]), OBTablette_l, OBCaption_l );
            end
            else if sClasseNom_l = 'TDBCheckBox' then
            begin
               nCompOngletNb_l := nCompOngletNb_l +
                     TraiteCheckBox( TDBCheckBox(Ecran.Components[nFICompNb_l]), OBCaption_l );
            end;
         end;
      end;
   end;

   SetControlVisible( 'PDIVERS' , (nCompOngletNb_l <> 0) );

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
function TOM_ZONESLIBRES.TraiteLabel( hLabel_p: THLabel; OBTablette_p, OBCaption_p : TOB ) : integer;
var
   wcFocus_l : TWinControl;
   sPrefixe_l, sTablette_l, sValLabel_l, sNomLabel_l : string;
   nTabInd_l : integer;
begin
   result := 0;

   sPrefixe_l := Copy( hLabel_p.Name, 1, Pos( '_', hLabel_p.Name ) - 1 );
   if sPrefixe_l <> TableToPrefixe( sTableParam_c ) then
   begin
      exit;
   end;

   wcFocus_l := hLabel_p.FocusControl;
   if wcFocus_l = nil then
   begin
      exit;
   end;

   if wcFocus_l.ClassName = 'THDBValComboBox' then
   begin
      sTablette_l := THDBValComboBox(wcFocus_l).DataType;
      nTabInd_l := StrToInt(sTablette_l[Length(sTablette_l)]) - 1;
      sValLabel_l := RechDom( sNomTablette_c, OBTablette_p.Detail[nTabInd_l].GetValue( TTToCode( sNomTablette_c)), false );
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

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 04/03/2003
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
function TOM_ZONESLIBRES.TraiteCheckBox( hCheckBox_p: TDBCheckBox; OBCaption_p : TOB ) : integer;
var
   sPrefixe_l, sValLabel_l, sNomLabel_l : string;
begin
   result := 0;

   sNomLabel_l := hCheckBox_p.Name;
   sPrefixe_l := Copy( hCheckBox_p.Name, 1, Pos( '_', hCheckBox_p.Name ) - 1 );
   if sPrefixe_l <> TableToPrefixe( sTableParam_c ) then
   begin
      sNomLabel_l := StringReplace( sNomLabel_l, sPrefixe_l, TableToPrefixe( sTableParam_c ), [rfIgnoreCase] );
      sNomLabel_l := StringReplace( sNomLabel_l, 'BOOL', 'COCHE', [rfIgnoreCase] );
   end;

   sValLabel_l := OBCaption_p.GetValue( sNomLabel_l );
   SetControlCaption( hCheckBox_p.Name, sValLabel_l );
   SetControlVisible( hCheckBox_p.Name, (sValLabel_l <> '') );

   if (sValLabel_l <> '') then
      result := 1;
end ;

////////////////////////////// INITIALIZATION //////////////////////////////

Initialization
   registerclasses ( [ TOM_ZONESLIBRES ] ) ;
end.


