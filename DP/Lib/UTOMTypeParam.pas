{***********UNITE*************************************************
Auteur  ...... : BM
Créé le ...... : 05/11/2002
Modifié le ... :   /  /
Description .. : Source TOM Générique pour les
                 TABLES : TOM_JUFORMEJUR, TOM_JUTYPECIVIL,
                 TOM_JUTYPEEVT, TOM_JUTYPEPER
Mots clefs ... : TOM;TYPEPARAM
*****************************************************************}

Unit UTOMTypeParam;
//////////////////////////////////////////////////////////////////
Interface
//////////////////////////////////////////////////////////////////
Uses
{$IFNDEF EAGLCLIENT}
     db,
{$ENDIF}
     UTOM, UTOB, HTB97;
//////////////////////////////////////////////////////////////////
Const csCar_f : char = '_';

//////////////////////////////////////////////////////////////////
// TOM de base
//////////////////////////////////////////////////////////////////
Type
   TOM_TYPEPARAM = Class (TOM)
      protected
         sCodeChamp_f, sTypeElement_f, sPrefixeChamp_c : string;
         bCodeModifiable_f, bCEG_c : boolean;

         function  IncCompteur( sTable_p, sChamp_p : string ) : integer;
         procedure FiltreActiver ( sFiltreListe_p : string ) ;
         function  TypeParamObligatoire( sCodeCont_p : string; var sTitre_p : string ) : boolean;
         function  TypeParamUtilise( sCodeChamp_p, sNomTable_p, sCodeCont_p : string; var sTitre_p : string ) : boolean;

      public
         procedure OnArgument ( sParams_p : String )   ; override;
         procedure OnNewRecord                ; override;
         procedure OnLoadRecord               ; override;
         procedure OnChangeField (F : TField) ; override;
         procedure OnUpdateRecord             ; override;
         procedure OnDeleteRecord             ; override;
         procedure OnAfterUpdateRecord        ; override ;
         procedure OnAfterDeleteRecord        ; override ;
         procedure OnClose                    ; override;

         procedure OnClick_BDuplique(Sender : TObject); virtual;

      private
         sPredefChamp_f, sVersionChamp_f : string;
         sNomTable_f, sFiltreListe_f : string;
         OBParam_f : TOB;

         function FormateCode( sCodeChamp_p : string ) : string;
   end ;

//////////////////////////////////////////////////////////////////
Implementation
//////////////////////////////////////////////////////////////////
uses
{$IFDEF EAGLCLIENT}
   eFichList,
{$ELSE}
   FichList,
{$ENDIF}
   hctrls, HMsgBox, SysUtils, Controls, dpJurOutils;
//////////////////////////////////////////////////////////////////
{*****************************************************************
Auteur ....... : BM
Date ......... : 05/10/02
Procédure .... : OnArgument
Description .. : Détermination du suffixe à utiliser
Paramètres ... : Liste des arguments
*****************************************************************}
procedure TOM_TYPEPARAM.OnArgument ( sParams_p : String ) ;
begin
   Inherited ;

   // $$$ JP 22/11/04 - pas les mêmes infos en eagl qu'en agl
{$IFDEF EAGLCLIENT}
   sNomTable_f     := TFFicheListe(Ecran).TableName;
   sPrefixeChamp_c := TableToPrefixe (sNomTable_f);
{$ELSE}
   sPrefixeChamp_c := TFFicheListe(Ecran).TableName;
   sNomTable_f     := PrefixeToTable (sPrefixeChamp_c);
{$ENDIF}

   sCodeChamp_f    := TFFicheListe(Ecran).CodeName;
   sPredefChamp_f  := sPrefixeChamp_c + '_PREDEFINI';
   sVersionChamp_f := sPrefixeChamp_c + '_VERSION';
   bCodeModifiable_f := true;

{$IFDEF EAGLCLIENT}
   sFiltreListe_f := TFFicheListe (Ecran).FRange;
{$ELSE}
   sFiltreListe_f := TFFicheListe (Ecran).Ta.Filter;
{$ENDIF}

   if GetControl( sPredefChamp_f ) <> nil then
      SetControlEnabled( sPredefChamp_f, false );
   if GetControl( sVersionChamp_f ) <> nil then
      SetControlEnabled( sVersionChamp_f, false );

   if GetControl('BDUPLIQUE') <> nil then
      TToolbarButton97(GetControl('BDUPLIQUE')).OnClick := OnClick_BDuplique;
end ;

{*****************************************************************
Auteur ....... : BM
Date ......... : 05/10/02
Procédure .... : OnNewRecord
Description .. : Force le champ PREDEFINI avec 'STD' pour que les
                 données ne soient pas écrasées au rechargement de
                 socref
Paramètres ... :
*****************************************************************}

procedure TOM_TYPEPARAM.OnNewRecord ;
begin
   Inherited ;
   if OBParam_f <> nil then
   begin
      OBParam_f.PutEcran( Ecran );
      FreeAndNil( OBParam_f );
   end;
   SetField( sPredefChamp_f, 'STD' );
   bCEG_c := false;
   SetField( sVersionChamp_f, 0 );
   SetControlEnabled('BDUPLIQUE', GetControlEnabled('BINSERT'));
end ;

{*****************************************************************
Auteur ....... : BM
Date ......... : 05/10/02
Procédure .... : OnLoadRecord
Description .. : Rend actif ou non le bouton "supprimer" en fonction
                 du champ '_PREDEFINI'
Paramètres ... :
*****************************************************************}

procedure TOM_TYPEPARAM.OnLoadRecord ;
begin
   Inherited ;
   bCEG_c := GetField( sPredefChamp_f ) = 'CEG';
   SetControlEnabled('BDELETE', not bCEG_c);
   SetControlEnabled('BDUPLIQUE', GetControlEnabled('BINSERT'));
end ;


{*****************************************************************
Auteur ....... : BM
Date ......... : 06/10/02
Procédure .... : OnChangeField
Description .. : En mode création, ajout d'un caractère spécial
                 au code
Paramètres ... :
*****************************************************************}

procedure TOM_TYPEPARAM.OnChangeField (F: TField);
{$ifndef GRC}
var
   sCodeCont_l, sCodeContOld_l, sTitre_l : string;
{$endif}
begin
   Inherited ;
   SetControlEnabled('BDUPLIQUE', GetControlEnabled('BINSERT'));
{$ifndef GRC}
   if (GetField(sPredefChamp_f) = 'CEG') and (DS.State = dsEdit) then
   begin
      sCodeCont_l := GetField(sCodeChamp_f);
      if TypeParamObligatoire(sCodeCont_l, sTitre_l) then
      begin
         if not bCodeModifiable_f then
         begin
            PGIInfo('Modification interdite, ce type ' + sTypeElement_f + ' est obligatoire.', sTitre_l );;
            TFFicheListe(Ecran).bDefaireClick( nil );
         end
         else
         begin
            if PGIAsk('Modification interdite, ce type ' + sTypeElement_f + ' est obligatoire.#13#10' +
                      'Voulez-vous le dupliquer?', sTitre_l ) = mrYes then
            begin
               if OBParam_f = nil then
               begin
                  OBParam_f := TOB.Create( sNomTable_f, nil, -1 );
                  OBParam_f.GetEcran( Ecran );
               end;
               // BM 17/06/05 : Fonctionnement différent pour les formes juridiques
               if sPrefixeChamp_c = 'JFJ' then
               begin
                  sCodeContOld_l := GetField (sCodeChamp_f) + '_';
                  OBParam_f.PutValue(sCodeChamp_f, sCodeContOld_l);
               end;
               OBParam_f.PutValue( F.FieldName, GetField(F.FieldName));
               DS.Cancel;
               TFFicheListe(Ecran).BInsertClick( nil );
            end
            else
            begin
               OBParam_f.Free;
               DS.Cancel;
               TFFicheListe(Ecran).bDefaireClick( nil );
            end;
         end;
      end;
   end
   else if (F.FieldName = sCodeChamp_f) then
   begin
      sCodeContOld_l := GetField (sCodeChamp_f);
      // BM 17/06/05 : Fonctionnement différent pour les formes juridiques
      if sPrefixeChamp_c = 'JFJ' then Exit;
      if (DS.State = dsInsert) and ( sCodeContOld_l <> '') then
      begin
         sCodeCont_l := FormateCode( sCodeChamp_f );
         if ( sCodeCont_l <> '' ) and ( sCodeContOld_l <> sCodeCont_l ) then
            SetField( sCodeChamp_f, sCodeCont_l );
      end;
   end
   else
{$endif GRC}
   if (F.FieldName = sPredefChamp_f) then
      bCEG_c := GetField( sPredefChamp_f ) = 'CEG';
end;

{*****************************************************************
Auteur ....... : BM
Date ......... : 05/10/02
Procédure .... : OnUpdateRecord
Description .. : En mode modification, ajout d'un caractère spécial
                 au code, raz de la version et passage en mode STD
Paramètres ... :
*****************************************************************}

procedure TOM_TYPEPARAM.OnUpdateRecord ;
var
   nVersion_l : integer;
begin
   Inherited ;
   if ( DS.State = dsInsert ) then
      SetControlEnabled( sCodeChamp_f, false );

   nVersion_l := GetField( sVersionChamp_f );
   nVersion_l := nVersion_l + 1;
   SetField( sVersionChamp_f, IntToStr(nVersion_l) );
   SetControlEnabled('BDUPLIQUE', GetControlEnabled('BINSERT'));
end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. MERIAUX
Créé le ...... : 21/07/2003
Modifié le ... : 21/07/2003
Description .. : controle des liens avec  la table JUTYPEEVT et Gestion
Suite ........ : des codes obligatoires
Mots clefs ... :
*****************************************************************}
procedure TOM_TYPEPARAM.OnDeleteRecord;
var
   sTitre_l, sCodeCont_l : string;
begin
   Inherited ;
   SetControlEnabled('BDUPLIQUE', GetControlEnabled('BINSERT'));
   sCodeCont_l := GetField( sCodeChamp_f );
   if TypeParamObligatoire( sCodeCont_l, sTitre_l ) then
   begin
      PGIBox('Suppression interdite, ce type ' + sTypeElement_f + ' est obligatoire.', sTitre_l );
      Lasterror := 1;
      Exit;
   end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. MERIAUX
Créé le ...... : 21/07/2003
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOM_TYPEPARAM.OnClose;
begin
   Inherited ;
   if OBParam_f <> nil then
      FreeAndNil( OBParam_f );
end ;

{*****************************************************************
Auteur ....... : BM
Date ......... : 26/11/02
Procédure .... : FormateCode
Description .. : Ajout d'un caractère spécial au code
Paramètres ... : Le code
*****************************************************************}
function TOM_TYPEPARAM.FormateCode( sCodeChamp_p : string ) : string;
Var
   sCodeCont_l : String;
   nCodeLg_l : integer;
begin
   sCodeCont_l := GetField( sCodeChamp_p );
   result := sCodeCont_l;
   if not bCodeModifiable_f then
      exit;

   nCodeLg_l := ChampToLength(sCodeChamp_p);

   if Pos( csCar_f, sCodeCont_l ) = 0 then
   begin
      sCodeCont_l := Copy(sCodeCont_l, 1, nCodeLg_l - 1 ) + csCar_f;
   end
   else if Pos( csCar_f, sCodeCont_l ) = Length(sCodeCont_l) then
      // Bonne position, ne fait rien
         sCodeCont_l := ''
   else if Pos( csCar_f, sCodeCont_l ) < Length(sCodeCont_l) then
   begin
      sCodeCont_l := StringReplace( sCodeCont_l, csCar_f, '', [rfReplaceAll, rfIgnoreCase]);
      sCodeCont_l := Copy(sCodeCont_l, 1, nCodeLg_l - 1 ) + csCar_f;
   end;
   result := sCodeCont_l;
end;

{*****************************************************************
Auteur ....... : BM
Date ......... : 06/10/02
Fonction ..... : IncCompteur
Description .. : Calcule un nouveau code interne disponible, type Integer
Paramètres ... : La table, le champ
Renvoie ...... : Le compteur incrémenté
*****************************************************************}

function TOM_TYPEPARAM.IncCompteur( sTable_p, sChamp_p : string ) : integer;
var
   nValeur_l : integer;
begin
   nValeur_l := MaxChampX(sTable_p, sChamp_p);
   result := nValeur_l + 1;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. MERIAUX
Créé le ...... : 21/07/2003
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
function TOM_TYPEPARAM.TypeParamObligatoire( sCodeCont_p : string; var sTitre_p : string ) : boolean;
var
   sLib_l : string;
begin
   sLib_l :=  RechDom( sNomTable_f, sCodeCont_p, False);
   sTitre_p := format('Type ' + sTypeElement_f + ' : %s - %s', [sCodeCont_p, sLib_l]);

   result := ExisteSQL('SELECT ' + sCodeChamp_f + ' ' +
                       'FROM '  + sNomTable_f + ' ' +
                       'WHERE ' + sCodeChamp_f + ' = "' + sCodeCont_p + '" ' +
                       '  AND ' + sPredefChamp_f + ' = "CEG"' );
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. MERIAUX
Créé le ...... : 21/07/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
function TOM_TYPEPARAM.TypeParamUtilise( sCodeChamp_p, sNomTable_p, sCodeCont_p : string; var sTitre_p : string ) : boolean;
var
   sLib_l : string;
begin
   sLib_l :=  RechDom( sNomTable_f, sCodeCont_p, False);
   sTitre_p := format('Type ' + sTypeElement_f + ' : %s - %s', [sCodeCont_p, sLib_l]);

   result := ExisteSQL('SELECT ' + sCodeChamp_p + ' ' +
                       'FROM '  + sNomTable_p + ' ' +
                       'WHERE ' + sCodeChamp_p + ' = "' + sCodeCont_p + '" ' );
end;
{***********A.G.L.Privé.*****************************************
Auteur  ...... : B.MERIAUX
Créé le ...... : 22/10/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
// $$$ JP 25/11/2004 - modif' pour gestion en eagl + construction filtre plus propre
procedure TOM_TYPEPARAM.FiltreActiver (sFiltreListe_p:string);
var
   strFiltre   :string;
begin
     if Trim (sFiltreListe_f) <> '' then
     begin
          strFiltre := Trim (sFiltreListe_f);
          if Trim (sFiltreListe_p) <> '' then
             strFiltre := strFiltre + ' AND ' + Trim (sFiltreListe_p);
     end
     else if Trim (sFiltreListe_p) <> '' then
          strFiltre := Trim (sFiltreListe_p);

     // Application du filtre à la liste
{$IFDEF EAGLCLIENT}
     strFiltre := Trim(StringReplace(strFiltre, '''', '"', [rfReplaceAll, rfIgnoreCase]));
     TFFicheListe(Ecran).SetNewRange('', strFiltre);
{$ELSE}
     TFFicheListe(Ecran).Ta.Filter   := strFiltre;
     TFFicheListe(Ecran).Ta.Filtered := TRUE;
{$ENDIF}

//   TFFicheListe (Ecran).Ta.Filter := sFiltreListe_f;
//   TFFicheListe (Ecran).Ta.Filtered := TRUE;
//   if TFFicheListe (Ecran).Ta.Filter <> '' then
//      TFFicheListe (Ecran).Ta.Filter := TFFicheListe (Ecran).Ta.Filter + ' AND ' + sFiltreListe_p
//   else
//      TFFicheListe (Ecran).Ta.Filter := sFiltreListe_p;
//   TFFicheListe (Ecran).Ta.Filtered := TRUE;}
end;

{***********A.G.L.***********************************************
Auteur  ...... : BM
Créé le ...... : 06/07/2005
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOM_TYPEPARAM.OnClick_BDuplique(Sender : TObject);
var
   sCodeContOld_l : string;
begin
   inherited;
   if OBParam_f = nil then
   begin
      OBParam_f := TOB.Create( sNomTable_f, nil, -1 );
      OBParam_f.GetEcran( Ecran );
   end;
   // BM 17/06/05 : Fonctionnement différent pour les formes juridiques
   if sPrefixeChamp_c = 'JFJ' then
   begin
      sCodeContOld_l := GetField (sCodeChamp_f) + '_';
      OBParam_f.PutValue(sCodeChamp_f, sCodeContOld_l);
   end;
//   OBParam_f.PutValue( F.FieldName, GetField(F.FieldName));
   DS.Cancel;
   TFFicheListe(Ecran).BInsertClick( nil );
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 06/07/2005
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOM_TYPEPARAM.OnAfterUpdateRecord ;
begin
   Inherited ;
   SetControlEnabled('BDUPLIQUE', GetControlEnabled('BINSERT'));
end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 06/07/2005
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOM_TYPEPARAM.OnAfterDeleteRecord ;
begin
   Inherited ;
   SetControlEnabled('BDUPLIQUE', GetControlEnabled('BINSERT'));
end ;

end.
