{***********UNITE*************************************************
Auteur  ...... : PL
Créé le ...... : 09/03/2001
Modifié le ... :   /  /
Description .. : Source TOF : AFTRADUCCHAMPLIBRE ()
Ancêtre permettant la traduction automatique des champs libres
Mots clefs ... : TOF;AFTRADUCCHAMPLIBRE
*****************************************************************}
Unit UTOFAFTRADUCCHAMPLIBRE ;

Interface

Uses StdCtrls, Controls, Classes,
{$IFDEF EAGLCLIENT}
{$ELSE}
     db,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ENDIF}
     forms, sysutils,
     ComCtrls,
     UtilPgi,
     EntGC,
  	 uTOFComm,
//uniquement en line
//   BTMUL_TOF,
     HCtrls, HEnt1, HMsgBox, UTOF,UtilGc,Ent1 ;

Type
  TOF_AFTRADUCCHAMPLIBRE = class(tTOFComm)
    procedure OnArgument (stArgument : String ) ; override ;
    Procedure TraiteLibreComboTB( champ,Souli : string; Nbre : integer);
    procedure OnLoad; override;
  end ;

Implementation


procedure TOF_AFTRADUCCHAMPLIBRE.OnArgument (stArgument : String ) ;
var     TLab : THLabel;
begin
  Inherited ;
  // on alimente le champ ISGESTAFF en fct gestion affaire ou GI
if ((Ecran.FindComponent ('ISGESTAFF'))<>nil) then
    begin
    If CtxSCot in V_PGI.PgiCOntexte then   SetControlChecked('ISGESTAFF',false)
      else   SetControlChecked('ISGESTAFF',True);
    end;
   // change les libelles champs libre + non affichage si pas gérer
    //clients
if ((Ecran.FindComponent ('YTC_TABLELIBRETIERS1'))<>nil) then
    begin
    GCMAJChampLibre (TForm (Ecran), False, 'COMBO', 'YTC_TABLELIBRETIERS', 10, '_');
    GCMAJChampLibre (TForm (Ecran), False, 'EDIT', 'YTC_VALLIBRE', 3, '_');
    GCMAJChampLibre (TForm (Ecran), False, 'EDIT', 'YTC_DATELIBRE', 3, '_');
    GCMAJChampLibre (TForm (Ecran), False, 'BOOL', 'YTC_BOOLLIBRE', 3, '_');
    GCMAJChampLibre (TForm (Ecran), False, 'EDIT', 'YTC_TEXTELIBRE', 3, '_');
    GCMAJChampLibre (TForm (Ecran), False, 'EDIT', 'YTC_RESSOURCE', 3, '_');
    end;
// Tache
if ((Ecran.FindComponent('ATA_LIBRETACHE1')) <> nil) then
  begin
    GCMAJChampLibre (TForm (Ecran), False, 'COMBO', 'ATA_LIBRETACHE', 10, '_');
    GCMAJChampLibre (TForm (Ecran), False, 'EDIT', 'ATA_DATELIBRE', 3, '_');
    GCMAJChampLibre (TForm (Ecran), False, 'EDIT', 'ATA_CHARLIBRE', 3, '_');
    GCMAJChampLibre (TForm (Ecran), False, 'EDIT', 'ATA_VALLIBRE', 3, '_');
    GCMAJChampLibre (TForm (Ecran), False, 'BOOL', 'ATA_BOOLLIBRE', 3, '_');
  end;
// AfPlanning
if ((Ecran.FindComponent('APL_LIBRETACHE1')) <> nil) then
  begin
    GCMAJChampLibre (TForm (Ecran), False, 'COMBO', 'APL_LIBRETACHE', 10, '_');
    GCMAJChampLibre (TForm (Ecran), False, 'EDIT', 'APL_DATELIBRE', 3, '_');
    GCMAJChampLibre (TForm (Ecran), False, 'EDIT', 'APL_CHARLIBRE', 3, '_');
    GCMAJChampLibre (TForm (Ecran), False, 'EDIT', 'APL_VALLIBRE', 3, '_');
    GCMAJChampLibre (TForm (Ecran), False, 'BOOL', 'APL_BOOLLIBRE', 3, '_');
  end;
                // affaires
if ((Ecran.FindComponent ('AFF_LIBREAFF1'))<>nil) then
    begin
    GCMAJChampLibre (TForm (Ecran), False, 'COMBO', 'AFF_LIBREAFF', 10, '_');
    GCMAJChampLibre (TForm (Ecran), False, 'EDIT', 'AFF_VALLIBRE', 3, '_');
    GCMAJChampLibre (TForm (Ecran), False, 'EDIT', 'AFF_DATELIBRE', 3, '_');
    GCMAJChampLibre (TForm (Ecran), False, 'EDIT', 'AFF_CHARLIBRE', 3, '_');
    GCMAJChampLibre (TForm (Ecran), False, 'BOOL', 'AFF_BOOLLIBRE', 3, '_');
    end;
if ((Ecran.FindComponent ('AFF_RESSOURCE1'))<>nil) then
    begin
    GCMAJChampLibre (TForm (Ecran), False, 'EDIT', 'AFF_RESSOURCE', 3, '_');
    end;
if ((Ecran.FindComponent ('EAF_RESSOURCE1'))<>nil) then
    begin // mcd 23/12/02
    GCMAJChampLibre (TForm (Ecran), False, 'EDIT', 'EAF_RESSOURCE', 3, '_');
    end;
                //ressource
if ((Ecran.FindComponent ('ARS_LIBRERES1'))<>nil) then
    begin     
    GCMAJChampLibre (TForm (Ecran), False, 'COMBO', 'ARS_LIBRERES', 10, '_');
    GCMAJChampLibre (TForm (Ecran), False, 'EDIT', 'ARS_VALLIBRE', 3, '_');
    GCMAJChampLibre (TForm (Ecran), False, 'EDIT', 'ARS_DATELIBRE', 3, '_');
    GCMAJChampLibre (TForm (Ecran), False, 'EDIT', 'ARS_CHARLIBRE', 3, '_');
    GCMAJChampLibre (TForm (Ecran), False, 'BOOL', 'ARS_BOOLLIBRE', 3, '_');
    end;
{$ifdef GCGC}
    //mcd 11/03/2005 ajout traitement équipe
if ((Ecran.FindComponent ('ARS_EQUIPERESS'))<>nil) then
    begin  //renseigne le champ avec les équipes associée à l'utilisatuer si confidentailité par planning
        // attention, pas de restriction en condition plus, car l'utilisateur à le droit de tout choisir quand même
        // si pas d'équipe associée à l'utilisateur, met l'équipe de la ressource
      ThMultiValComboBox(GetControl('ARS_EQUIPERESS')).text := VH_GC.AfEquipe;
    end;
{$endif}
                //article
if ((Ecran.FindComponent ('GA_LIBREART1'))<>nil) then
    begin
    GCMAJChampLibre (TForm (Ecran), False, 'COMBO', 'GA_LIBREART', 10, '_');
    GCMAJChampLibre (TForm (Ecran), False, 'EDIT', 'GA_VALLIBRE', 3, '_');
    GCMAJChampLibre (TForm (Ecran), False, 'EDIT', 'GA_DATELIBRE', 3, '_');
    GCMAJChampLibre (TForm (Ecran), False, 'EDIT', 'GA_CHARLIBRE', 3, '_');
    GCMAJChampLibre (TForm (Ecran), False, 'BOOL', 'GA_BOOLLIBRE', 3, '_');
    end;
if ((Ecran.FindComponent ('GL_LIBREART1'))<>nil) then
    begin  //mcd 12/09/03 oubli ..
    GCMAJChampLibre (TForm (Ecran), False, 'COMBO', 'GL_LIBREART', 10, '_');
    end;

// Compétences
if ((Ecran.FindComponent ('PCO_TABLELIBRERH1'))<>nil) then
    begin
    GCMAJChampLibre (TForm (Ecran), False, 'COMBO', 'PCO_TABLELIBRERH', 5, '_');
    end;
// Compétences/ article
if ((Ecran.FindComponent ('AAC_TABLELIBRECA1'))<>nil) then
    begin
    GCMAJChampLibre (TForm (Ecran), False, 'COMBO', 'AAC_TABLELIBRECA', 3, '_');
    end;
// Compétences/ ressource
if ((Ecran.FindComponent ('PCH_TABLELIBRECR1'))<>nil) then
    begin
    GCMAJChampLibre (TForm (Ecran), False, 'COMBO', 'PCH_TABLELIBRECR', 3, '_');
    end;
{$IFDEF GCGC}
TLab:=THLabel (Ecran.FindComponent ('TGA_FAMILLENIV1'));
if tlab <>NIL then
  ChangeLibre2('TGA_FAMILLENIV1',Ecran); //AB-20050111
TLab:=THLabel (Ecran.FindComponent ('TGA_FAMILLENIV2'));
if tlab <>NIL then
  ChangeLibre2('TGA_FAMILLENIV2',Ecran);
TLab:=THLabel (Ecran.FindComponent ('TGA_FAMILLENIV3'));
if tlab <>NIL then
  ChangeLibre2('TGA_FAMILLENIV3',Ecran);
   //ligne
TLab:=THLabel (Ecran.FindComponent ('TGL_FAMILLENIV1'));
if tlab <>NIL then
  ChangeLibre2('TGL_FAMILLENIV1',Ecran); //AB-20050111
TLab:=THLabel (Ecran.FindComponent ('TGL_FAMILLENIV2'));
if tlab <>NIL then
  ChangeLibre2('TGL_FAMILLENIV2',Ecran);
TLab:=THLabel (Ecran.FindComponent ('TGL_FAMILLENIV3'));
if tlab <>NIL then
  ChangeLibre2('TGL_FAMILLENIV3',Ecran);
{$ELSE}
TLab:=THLabel (Ecran.FindComponent ('TGA_FAMILLENIV1'));
if tlab <>NIL then SetCOntrolText('TGA_FAMILLENIV1', RechDom('GCLIBFAMILLE','LF1',false));
TLab:=THLabel (Ecran.FindComponent ('TGA_FAMILLENIV2'));
if tlab <>NIL then SetCOntrolText('TGA_FAMILLENIV2', RechDom('GCLIBFAMILLE','LF2',false));
TLab:=THLabel (Ecran.FindComponent ('TGA_FAMILLENIV3'));
if tlab <>NIL then SetCOntrolText('TGA_FAMILLENIV3', RechDom('GCLIBFAMILLE','LF3',false));
   //ligne
TLab:=THLabel (Ecran.FindComponent ('TGL_FAMILLENIV1'));
if tlab <>NIL then SetCOntrolText('TGL_FAMILLENIV1', RechDom('GCLIBFAMILLE','LF1',false));
TLab:=THLabel (Ecran.FindComponent ('TGL_FAMILLENIV2'));
if tlab <>NIL then SetCOntrolText('TGL_FAMILLENIV2', RechDom('GCLIBFAMILLE','LF2',false));
TLab:=THLabel (Ecran.FindComponent ('TGL_FAMILLENIV3'));
if tlab <>NIL then SetCOntrolText('TGL_FAMILLENIV3', RechDom('GCLIBFAMILLE','LF3',false));
{$ENDIF GCGC}

// piece
if ((Ecran.FindComponent ('GP_LIBREPIECE1'))<>nil) then
    GCMAJChampLibre (TForm (Ecran), False, 'COMBO', 'GP_LIBREPIECE', 3, '_');
if ((Ecran.FindComponent ('GP_DATELIBRE1'))<>nil) then
    GCMAJChampLibre (TForm (Ecran), False, 'COMBO', 'GP_DATELIBRE', 3, '_');
if ((Ecran.FindComponent ('GP_LIBRETIERS1'))<>nil) then
    GCMAJChampLibre (TForm (Ecran), False, 'COMBO', 'GP_LIBRETIERS', 10, '_');
if ((Ecran.FindComponent ('GP_LIBREAFF1'))<>nil) then
    GCMAJChampLibre (TForm (Ecran), False, 'COMBO', 'GP_LIBREAFF', 3, '_');
// pas gérer pour l'instant en titre libre (10/04/01) if ((Ecran.FindComponent ('GP_LIBREPIECE1'))<>nil) then
//     GCMAJChampLibre (TForm (Ecran), False, 'COMBO', 'GP_LIBREPIECE', 3, '_');
if ((Ecran.FindComponent ('GP_TIERSSAL1'))<>nil) then
    GCMAJChampLibre (TForm (Ecran), False, 'COMBO', 'GP_TIERSSAL', 3, '_');

//******** spécificité pour le TB ***********/
if (ecran.Name= 'AFTABLEAUBORD') or (ecran.Name= 'AFTBVIEWERMULTI') or (ecran.Name= 'AFBALANCE') or
   (ecran.Name= 'AFGRANDLIVRE') or (Ecran.name='AFAPPREC_ETAT') or (ecran.Name= 'AFETAT_CTRL_PROD') then
   begin
   // tiers
   TraiteLibreComboTB( 'LIBRETIERS','',10);
   TLab:=THLabel (Ecran.FindComponent ('TATB_RES1TIERS'));
   if tlab <>NIL then begin
     SetCOntrolText('TATB_RES1TIERS', RechDom('GCZONELIBRETIE','CR1',false));
     if (GetControlText('TATB_RES1TIERS') = '.-')   then
         begin
         SetCOntrolVisible('TATB_RES1TIERS',False); SetCOntrolVisible('TATB_RES1TIERS_',False);
         SetCOntrolVisible('ATBXRES1TIERS',False);SetCOntrolVisible('ATBXRES1TIERS_',False);
         end;
     end;
   TLab:=THLabel (Ecran.FindComponent ('TATB_RES2TIERS'));
   if tlab <>NIL then begin
      SetCOntrolText('TATB_RES2TIERS', RechDom('GCZONELIBRETIE','CR2',false));
        if (GetControlText('TATB_RES2TIERS') = '.-')   then begin
        SetCOntrolVisible('TATB_RES2TIERS',False); SetCOntrolVisible('TATB_RES2TIERS_',False);
        SetCOntrolVisible('ATBXRES2TIERS',False);  SetCOntrolVisible('ATBXRES2TIERS_',False);
        end;
      end;
   TLab:=THLabel (Ecran.FindComponent ('TATB_RES3TIERS'));
   if tlab <>NIL then begin
      SetCOntrolText('TATB_RES3TIERS', RechDom('GCZONELIBRETIE','CR3',false));
     if (GetControlText('TATB_RES3TIERS') = '.-')   then begin
        SetCOntrolVisible('TATB_RES3TIERS',False); SetCOntrolVisible('TATB_RES3TIERS_',False);
        SetCOntrolVisible('ATBXRES3TIERS',False);  SetCOntrolVisible('ATBXRES3TIERS_',False);
        end;
      end;
   // Affaire
   TraiteLibreComboTB( 'LIBREAFF','',10);
   TraiteLibreComboTB( 'AFFRESSOURCE','_',3);     // 15/10/01 ajout '-'
   // articles
   TLab:=THLabel (Ecran.FindComponent ('TATB_FAMILLENIV1'));
   if tlab <>NIL then SetCOntrolText('TATB_FAMILLENIV1', RechDom('GCLIBFAMILLE','LF1',false));
   TLab:=THLabel (Ecran.FindComponent ('TATB_FAMILLENIV2'));
   if tlab <>NIL then SetCOntrolText('TATB_FAMILLENIV2', RechDom('GCLIBFAMILLE','LF2',false));
   TLab:=THLabel (Ecran.FindComponent ('TATB_FAMILLENIV3'));
   if tlab <>NIL then SetCOntrolText('TATB_FAMILLENIV3', RechDom('GCLIBFAMILLE','LF3',false));
   end;
   //AB-200510- Multi sociétés
   if assigned (GetControl ('MULTIDOSSIER')) and EstBaseMultiSoc then
   begin
      THValComboBox(GetControl('MULTIDOSSIER')).Value := MS_CODEREGROUPEMENT;
   end;
    // on regarde si le user est bloqué au niveau établissement et on le gère
PositionneEtabUser(GetControl('ARS_ETABLISSEMENT')) ;
PositionneEtabUser(GetControl('AFF_ETABLISSEMENT')) ;
PositionneEtabUser(GetControl('GCL_ETABLISSEMENT')) ;
PositionneEtabUser(GetControl('ATBXETABLISSEMENT')) ;
PositionneEtabUser(GetControl('GP_ETABLISSEMENT')) ;
PositionneEtabUser(GetControl('GL_ETABLISSEMENT')) ; 
end ;

procedure TOF_AFTRADUCCHAMPLIBRE.OnLoad ;
begin
  Inherited ;
end ;

Procedure TOF_AFTRADUCCHAMPLIBRE.TraiteLibreComboTB( champ ,souli: string; Nbre : integer);
Var TLab : THLabel ;
    i : integer;
    Num,PrefixeTT,Lib : string;
begin
PrefixeTT:= '';
if Champ = 'LIBRETIERS'   then PrefixeTT := 'CT';
if Champ = 'LIBREAFF'     then PrefixeTT := 'MT';
if Champ = 'AFFRESSOURCE' then PrefixeTT := 'MR';

For i:=1 to Nbre do
   begin
   if (i < 10) then Num := IntToStr(i) else Num := 'A';
   TLab:=THLabel (Ecran.FindComponent ('TATB_' + Champ+ Num));
   if TLab <>NIL then
      begin
      Lib := RechDomZoneLibre (PrefixeTT+Num ,false);
      SetControlText('TATB_'+Champ+Num , lib);
      if (Lib = '.-')   then
         begin
         SetCOntrolVisible('TATB_'+ Champ+ Num,False); SetCOntrolVisible('TATB_'+ Champ+ Num,False);
         SetCOntrolVisible('ATBX' + Champ+ Num,False); SetCOntrolVisible('ATBX' + Champ+ Num,False);
         end;
      if (souli <>'') then begin
         TLab:=THLabel (Ecran.FindComponent ('TATB_' + Champ+ Num+Souli));
         if TLab <>NIL then
            begin
            Lib := RechDomZoneLibre (PrefixeTT+Num ,false);
            if (Lib = '.-')   then
               begin
               SetCOntrolVisible('TATB_'+ Champ+ Num+souli,False); SetCOntrolVisible('TATB_'+ Champ+ Num+souli,False);
               SetCOntrolVisible('ATBX' + Champ+ Num+souli,False); SetCOntrolVisible('ATBX' + Champ+ Num+souli,False);
               end;
            end;
         end;
      end;
   end;
end;

Initialization
  registerclasses ( [ TOF_AFTRADUCCHAMPLIBRE ] ) ;
end.
