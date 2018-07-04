unit dpOutils;

interface

uses Forms,SysUtils,HEnt1, Db, controls,HCtrls, Classes,
     Windows, UTOB, UTOM, stdctrls, comctrls, hmsgbox, MailOL,
{$IFDEF VER150}
     Variants,
{$ENDIF}
{$IFDEF EAGLCLIENT}
     MaineAGL, eMul, eFiche,
{$ELSE}
     HDB,
     {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
     Fe_Main, Mul, Fiche,
{$ENDIF}
     PGIAppli, HStatus, galOutil, UTof, HPanel, uiUtil;

function NbEnreg(table, champ, crit: String): Integer;

// function  IncrementeSeqNo(sTable, sNoSeq, sCle : String; No : Integer): Integer;
function IncrementeSeqNoOrdre(prefixe, table : String; GuidPer : String): Integer;
procedure MajNbFilialesOrga(GuidPer : String);
procedure Formate00(var chp : TField);
procedure SelToutSurEntree(Sender : TObject);
function  FctDesPropDeTitres(Formjur : String; sFonction_p : string = '') : string;
function  RecupNomAnnuLien(guidperdos, fonction : String; sTypePerDos_p : string = ''): String;
function  GetDateCreation (guidper : String) : TDateTime;
function  SupprimeAttach  (guidper: String; Fnc, Fnc2: String): Integer;
procedure LanceFicheDP(Nature, Fiche, Range, Lequel, Argument, table, prefix: String; Inside:THPanel=nil);//LMO(param)

procedure UpdateParamSocDossier(NomParam, Valeur: String);
procedure AfficheLibTablesLibres(LaTof: TOF);
function  GuidPerToCodeTiers(sGuidPer: String): String;
function  IsPersonneCabinet(sGuidPer_p : String): Boolean;
function  GetDirigeant(sGuidPer, sFormeJur : String; var sFctGrt, sFctLib, sNomDirig: String): String;
function  IsTabEmpty (Tab:TTabSheet):boolean;

procedure ReloadTomInsideAfterInsert(MyFiche_p : TFFiche; DS_P : TDataSet;
                                     aosKeys_p : array of string; aovCode_p : array of variant);

procedure DPLibelleLibres;

// $$$ JP 17/10/06: identifiant base commune (GUID), créer si pas encore défini
function dpGetBaseGuid:string;

// $$$ JP 26/10/06: blob en string
function BlobToString (Texte:string):string;

function FaitLeSelect (Ecran:TForm): string ; //LM20070404
procedure GroupBoxEnabled (frm:TForm; gBox:TControl; Sauf : string; ok:Boolean) ;//LM20070511

function  ControleCleFRP( sNoAdmin_p, sNoAdhesion_p, sNoDhCompl_p : string ) : boolean;


/////////// IMPLEMENTATION ////////////
implementation

uses
   {$IFDEF DP}
   DPJUROutilsDossiers,
   {$ENDIF DP}
  UDossierSelect,
  AnnOutils, DpJurOutils, extctrls, galSystem,
  paramsoc,
  EventDecla
   ,CbpMCD
   ,CbpEnumerator
  ;

// supression d'un enregistrement




function NbEnreg(table, champ, crit: String): Integer;
// Bien préciser le critère complet (avec WHERE)
var Q : TQuery;
begin
  Result := 0;
  Q := OpenSQL('SELECT count('+champ+') FROM '+table+' '+crit, True);
  if (Q<>nil) and (Not Q.EOF) and (Not VarIsNull(Q.Fields[0].Value)) then
    Result := Q.Fields[0].AsInteger;
  Ferme(Q);
end;


//------------------------------------------------------------------------------
//---- Mise à jour des élts DPORGA liés à la présence de filiales
//------------------------------------------------------------------------------
procedure MajNbFilialesOrga(GuidPer : String);
var Nb: Integer;
    Coche : String;
    TobDPOrga : TOB;
begin
  if GuidPer='' then exit;

  // Filiales (attention : typedos = 'DP')
  Nb := NbEnreg('ANNULIEN', '*', 'where ANL_FONCTION="FIL" AND ANL_GUIDPERDOS="'
                +GuidPer+ '" AND ANL_TYPEDOS="DP" AND ANL_NOORDRE=1');
  Coche := '-';
  if Nb>0 then Coche := 'X';

  // Création de l'enregistrement si nécessaire
  if Not ExisteSQL('select 1 from DPORGA where DOR_GUIDPER="'+GuidPer+'"') then
  begin
     // #### n'utilise pas la TOM DPORGA
     TobDPOrga := TOB.Create('DPORGA', Nil, -1);
     TobDPOrga.InitValeurs;
     TobDPOrga.PutValue ('DOR_GUIDPER', GuidPer);
     TobDPOrga.PutValue ('DOR_NBRATTACH', IntToStr(Nb) );
     TobDPOrga.PutValue ('DOR_ATTACHEMENT', Coche );
     TobDPOrga.InsertDB (nil);
     TobDPOrga.Free;
  end
  else
     ExecuteSQL('update DPORGA set DOR_NBRATTACH='+IntToStr(Nb)+','
               +' DOR_ATTACHEMENT="'+Coche+'" where DOR_GUIDPER="'+GuidPer+'"');
end;


//-----------------------------------------------------------------------------
// fonction qui incrémente en séquence une clé composée de NO + no de sequence
//-----------------------------------------------------------------------------
{ #### MD 20/02/06 : inutile et mal foutue ?
function IncrementeSeqNo(sTable, sNoSeq, sCle : String; No : Integer): Integer;
var
   Q: TQuery;
   sWhere: String;
begin
  sWhere := '';
  if sCle <> '' then sWhere := ' where '+sCle+ '= ' + (InttoStr (No));
  Q := OpenSQL('select max('+ sNoSeq +') as maxno from ' + sTable + sWhere, TRUE);
  if not Q.eof then
    if not VarIsNull(Q.Fields[0].Value) then
      Result := Q.Fields[0].AsInteger+1
    else
      Result := 1
  else
    Result := 1;
  Ferme(Q);
end;
}

//-----------------------------------------------------------------------------
// sur sortie de chaque zone, tronque les chiffres saisis
// au delà du 2ème après la virgule, ou met 0 si zone vide
//-----------------------------------------------------------------------------
procedure Formate00(var chp : TField);
begin
  if VarIsNull(chp.AsVariant) then
    chp.value := 0
  else
    chp.Value := FormatFloat('##0.00', chp.value);
end;


//------------------------------------------------------------------------------
// ---- sélection dès l'entrée du contenu de la zone
//------------------------------------------------------------------------------
procedure SelToutSurEntree(Sender : TObject);
// sur entrée dans chaque zone, présélectionne tout le contenu
begin
  if Sender is TEdit then
    begin
    TEdit(Sender).SelStart := 0;
    TEdit(Sender).SelLength := Length(TEdit(Sender).text);
    end
{$IFDEF EAGLCLIENT}
{$ELSE}
  else if Sender is THDBEdit then
    begin
    THDBEdit(Sender).SelStart := 0;
    THDBEdit(Sender).SelLength := Length(THDBEdit(Sender).text);
    end
{$ENDIF}
  else if Sender is THCritMaskEdit then
    begin
    THCritMaskEdit(Sender).SelStart := 0;
    THCritMaskEdit(Sender).SelLength := Length(THCritMaskEdit(Sender).text);
    end;

    //#### plus tard, tester les autres types de contrôles
end;


//------------------------------------------------------------------------------
//---- Incrémente le No ordre de l'intervention pour ce dossier pour ce type
//---- de liaison (à affiner sans doute en fonction du sens attribué au no ordre)
//------------------------------------------------------------------------------
function IncrementeSeqNoOrdre(prefixe, table : String; GuidPer : String): Integer;
var
   sOrdre : string;
   sCle   : string;
   Q: TQuery;
begin
  Result := 1;
  sOrdre := prefixe + '_NOORDRE';
  sCle   := prefixe + '_GUIDPER';
  //select max (DIJ_NOORDRE) from DPNTVJUR where DIJ_GUIDPER = 1
  Q := OpenSQL('select max ('+ sOrdre + ') as maxno from '+ table
   + ' where '+ sCle + ' = "' + GuidPer + '"', TRUE);
  if not q.eof then
    if not VarIsNull(Q.Fields[0].Value) then
      Result := (Q.Fields[0].AsInteger+1);

  Ferme(Q);
end;


function GetDateCreation (guidper : String) : TDateTime;
var
   Q : TQuery;
begin
  Q := OpenSQL('select JUR_DEBACT from JURIDIQUE where JUR_GUIDPERDOS="' + GuidPer + '" AND '
   + 'JUR_TYPEDOS="STE" AND JUR_NOORDRE=1', TRUE) ;
  if not Q.Eof then
    Result := StrToDate (Q.FindField('JUR_DEBACT').AsString)
  else
    Result := iDate1900;
  Ferme (Q);
end;


function FctDesPropDeTitres(Formjur : String; sFonction_p : string = '') : string;
// liste la (ou les) fonctions possibles des propriétaires de titres
// la plupart du temps, lorsqu'il y en a, c'est ASS ou ACT.
// mais pas toujours (voir les stés en commandité simple ou par action...)
var sRequete : String;
    Q        : TQuery;
begin
  Result := '';
  sRequete := 'select JTF_FONCTION '
   +'from (JUTYPEFONCT left join JUFONCTION on JFT_FONCTION=JTF_FONCTION) '
   +'where (JFT_FORME="'+Formjur+'" AND JFT_TYPEDOS="STE" AND JFT_TIERS<> "X") '
   + 'AND (JTF_TITRE="X") ';

   //$$$ BM : 12/08/2004 : condition supplémentaire
   if sFonction_p <> '' then
      sRequete := sRequete + 'AND JFT_FONCTION = "' + sFonction_p + '" ';
   sRequete := sRequete + 'ORDER BY JFT_TRI';

  Q := OpenSQL(sRequete,true);
  if not Q.Eof then
    Result := Q.FindField('JTF_FONCTION').AsString;
  Ferme(Q);
end;


function RecupNomAnnuLien(guidperdos, fonction : String; sTypePerDos_p : string = ''): String;
var Q : TQuery;
begin
  Result := '';
  if sTypePerDos_p <> '' then
    sTypePerDos_p := ' AND ANL_TYPEPERDOS = "' + sTypePerDos_p + '"';

  Q :=  OpenSQL('SELECT ANL_NOMPER FROM ANNULIEN WHERE ANL_GUIDPERDOS="'+guidperdos+'"'
   +' AND ANL_FONCTION="'+fonction+'"' + sTypePerDos_p, True);
  if not Q.eof then
    Result := Q.Fields[0].AsString;
  Ferme(Q);
end;


function SupprimeAttach(Guidper: String; Fnc, Fnc2: String): Integer;
var Msg : String;
begin
  Result := 1;
  if Fnc='' then exit; // on ne supprime pas les liens toutes fcts confondues !

  Msg := 'Cette opération entraînera la suppression de tous les liens correspondants. Voulez-vous supprimer ?';
  if PGIAsk(Msg, TitreHalley)=mrYes then
    begin
    ExecuteSQL('delete from ANNULIEN where ANL_GUIDPERDOS = "'
      + GuidPer + '" and ANL_FONCTION="'+Fnc+'"');
    if Fnc2 <> '' then
      ExecuteSQL('delete from ANNULIEN where ANL_GUIDPERDOS = "'
       + GuidPer + '" and ANL_FONCTION="'+Fnc2+'"');
    Result := 0; // plus de liens => succès
    end;
end;


procedure LanceFicheDP(Nature, Fiche, Range, Lequel, Argument, table, prefix: String; Inside:THPanel=nil);    //LMO(param)
// vérifie l'existence d'un enregistrement basé sur la clé _GUIDPER (Lequel) - ou étendue pour JUR-
// puis lance la fiche en modification
// (évite perte de l'enreg qd on modifie/valide 2 fois sur un nouvel enreg)
var Q : TQuery;
    arg, action, typdos, typfich, synthes, guidperdos, SQL,
    sNoOrdre_l, sForme_l, sCodeDos_l : String;
    Annee, Mois, Jour : Word;
    ChMois,MoisClos : String;
    TobDpFiscal : TOB;
    TomDpFiscal : TOM;
begin
   if Lequel='' then exit;

   //--- Construction / éxécution requête sur l'enregistrement DP ---
   // DP juridique
   if (table = 'JURIDIQUE') then
   begin
       arg := Argument;
       action := ReadTokenSt(arg);

       // BM 26/01/2006 : fusion Fiches et TOM juridiques
       guidperdos := ReadTokenSt(arg);
       if guidperdos='' then exit; // sortie si pas de clé
       typdos := ReadTokenSt(arg); // pas toujours STE ?
       if typdos<>'STE' then typdos := 'STE';

       typfich := ReadTokenSt(arg);
       synthes := ReadTokenSt(arg);

       SQL := 'SELECT 1 FROM JURIDIQUE WHERE JUR_GUIDPERDOS="'+guidperdos+'"'
            + ' and JUR_TYPEDOS="STE" and JUR_NOORDRE=1'
   end
   // Liens juridiques
   else if (table = 'ANNULIEN') and (Fiche = 'ANNULIEN_TRI') then
   begin
       guidperdos := ReadTokenSt(Lequel);
       if guidperdos = '' then exit; // sortie si pas de clé
       typdos := ReadTokenSt(Lequel);
       if typdos <> 'STE' then typdos := 'STE';
       sNoOrdre_l := ReadTokenSt(Lequel); // Toujours à 1
       typfich := Argument;

       SQL := 'SELECT 1 FROM JURIDIQUE WHERE JUR_GUIDPERDOS = "' + guidperdos +'"'
        + '  AND JUR_TYPEDOS = "' + typdos + '" AND JUR_NOORDRE = ' + sNoOrdre_l;
   end
   // Autres tables DP
   else {if (table = 'DPORGA') or (table = 'DPFISCAL') or (table = 'DPSOCIAL') then}
   begin
       SQL := 'SELECT 1 FROM '+table+' WHERE '+prefix+'_GUIDPER="'+Lequel+'"';
   end;

   //--- Si absence d'enregistrement, il faut passer en mode CREATION ---
   if Not ExisteSQL(SQL) then
   begin
       if (table='JURIDIQUE') then
       begin
         {$IFDEF DP}
         DPJURCreationDossierJuridique(guidperdos, typdos, 1);
         {$ENDIF DP}
       end

       // Liens juridiques
       else if (table = 'ANNULIEN') and (Fiche = 'ANNULIEN_TRI') then
       begin
           // si pas d'enregistrement juridique, pas de liens possibles
           // donc on ne lance pas la fiche
           exit;
       end
       else if (table='DPFISCAL') then
       begin
           // FQ 11185 & 11698
           TobDpFiscal := TOB.Create('DPFISCAL', Nil, -1);
           if not TobDpFiscal.SelectDB ('"' + Lequel + '"', nil, TRUE) then
           begin
                TomDpFiscal := CreateTom('DPFISCAL', Nil, True, True);
                TobDpFiscal.InitValeurs;
                TobDpFiscal.PutValue ('DFI_GUIDPER', Lequel);
                TomDpFiscal.InitTOB(TobDpFiscal);
                // FQ 11768 : pourquoi InitTob vide DFI_GUIDPER ?!!
                TobDpFiscal.PutValue ('DFI_GUIDPER', Lequel);
                TobDpFiscal.InsertOrUpdateDB(False);
                FreeAndNil(TomDpFiscal);
           end;
           FreeAndNil(TobDpFiscal);
       end
       // Autres tables DP
       else
       begin
           {
           //MD 05/06/06 - C'était une création sauvage, avec en plus mauvaise clé !
           // Q.Insert;
           // InitNew(Q);
           // Q.FindField(prefix+'_GUIDPER').AsString := AglGetGuid(); !!!
           // => il fallait : Q.FindField(prefix+'_GUIDPER').AsString := Lequel;
           // Q.Post; }
           Argument := StringReplace(Argument, 'ACTION=MODIFICATION;', 'ACTION=CREATION;', []);
           // Pour les appels sans le ; final ocazou!
           Argument := StringReplace(Argument, 'ACTION=MODIFICATION', 'ACTION=CREATION', []);
       end;
   end

   //--- Si l'enreg existe ---
   else
   begin
     // Complète les arguments pour les liens juridiques
     if (table = 'ANNULIEN') and (Fiche = 'ANNULIEN_TRI') then
     begin
         Q := OpenSQL('SELECT JUR_FORME, JUR_CODEDOS FROM JURIDIQUE WHERE JUR_GUIDPERDOS = "' + guidperdos +'"'
                    + '  AND JUR_TYPEDOS = "' + typdos + '" AND JUR_NOORDRE = ' + sNoOrdre_l, True);
         sForme_l := Q.FindField('JUR_FORME').AsString;
         sCodeDos_l := Q.FindField('JUR_CODEDOS').AsString;
         Argument := guidperdos + ';' + typdos + ';' + sNoOrdre_l + ';' +
                     sForme_l + ';' + sCodeDos_l + ';' + typfich;
         Ferme(Q);
     end

     // $$$JP 17/10/2003: màj des dates si existantes dans la table DPTABCOMPTA
     //+LM20070412 else if (table = 'DPORGA') and (Lequel <> '') then
     {+GHA - 09.01.2008
      Modif :
        * Argument Lequel remplacé par guidperdos car Lequel est initialisé avec une valeur GuidPer;STE;1
        * 1er ordre Sql update inversion JUR_DATEFINEX=DTC_DATEDEB, JUR_DATEDEBUTEX=DTC_DATEFIN
                       corrigé par JUR_DATEFINEX=DTC_DATEFIN, JUR_DATEDEBUTEX=DTC_DATEDEB
     -GHA - }
     else if (table = 'JURIDIQUE') then //and (Lequel <> '') then   //-LM20070412 LM>MD Ok ? Ne semble plus d'actualité
     begin
         // $$$JP 22/10/2003: mise à jour durée exercice précédent avec date compta "N-"
         ExecuteSQL (//+LM20070412 'UPDATE DPORGA SET DOR_DATEDEBUTEX=DTC_DATEDEB,DOR_DATEFINEX=DTC_DATEFIN,DOR_DUREE=DTC_DUREE ' +
                     //'update JURIDIQUE set JUR_DATEFINEX=DTC_DATEDEB, JUR_DATEDEBUTEX=DTC_DATEFIN, JUR_DUREEEX=DTC_DUREE ' + //-LM20070412 LM>MD A confirmer
                     'update JURIDIQUE set JUR_DATEFINEX=DTC_DATEFIN, JUR_DATEDEBUTEX=DTC_DATEDEB, JUR_DUREEEX=DTC_DUREE ' + //-GHA20080109 inversion date fin et date deb
                     'FROM DOSSIER,DPTABCOMPTA ' +
                     'where JUR_GUIDPERDOS="' + guidperdos + '" AND JUR_GUIDPERDOS=DOS_GUIDPER AND DOS_NODOSSIER=DTC_NODOSSIER AND DTC_MILLESIME="N"');//LM20070412 'WHERE DOR_GUIDPER="' + Lequel + '" AND DOR_GUIDPER=DOS_GUIDPER AND DOS_NODOSSIER=DTC_NODOSSIER AND DTC_MILLESIME="N"');

         ExecuteSQL ('update JURIDIQUE set JUR_DUREEEXPREC=DTC_DUREE ' + //LM20070412 UPDATE DPORGA SET DOR_DUREEPREC=DTC_DUREE ' +
                     'FROM DOSSIER,DPTABCOMPTA ' +
                     'where JUR_GUIDPERDOS="' + guidperdos + '" AND JUR_GUIDPERDOS=DOS_GUIDPER AND DOS_NODOSSIER=DTC_NODOSSIER AND DTC_MILLESIME="N-"');//LM20070412 'WHERE DOR_GUIDPER="' + Lequel + '" AND DOR_GUIDPER=DOS_GUIDPER AND DOS_NODOSSIER=DTC_NODOSSIER AND DTC_MILLESIME="N-"');

         // MD 24/06/2005: mise à jour ANN_MOISCLOTURE, même principe que DOR_DATEFINEX
         // mais ce serait mieux fait en temps réel par la compta (fonctions de CTableauBord)
         ChMois := '00';
         Q := OpenSQL('select JUR_DATEFINEX,ANN_MOISCLOTURE from JURIDIQUE ' +
                      'left join annuaire on jur_guidperdos = ann_guidper ' +
                      'where JUR_GUIDPERDOS="' + guidperdos + '"', True); //LM20070412 Q := OpenSQL('SELECT DOR_DATEFINEX FROM DPORGA WHERE DOR_GUIDPER="' + Lequel + '"', True);
         if (Not Q.Eof)
          and (Not VarIsNull(Q.FindField('JUR_DATEFINEX').Value)) //LM20070412 and (Not VarIsNull(Q.FindField('DOR_DATEFINEX').Value))
           and (Q.FindField('JUR_DATEFINEX').AsDateTime <> iDate1900) then //LM20070412 and (Q.FindField('DOR_DATEFINEX').AsDateTime <> iDate1900) then
         begin
             DecodeDate(Q.FindField('JUR_DATEFINEX').AsDateTime, Annee, Mois, Jour);
             ChMois:=Format ('%2.2d',[Mois]);
             MoisClos := Q.FindField('ANN_MOISCLOTURE').AsString;
         end;
         Ferme(Q);
         if (ChMois <> '00') and (MoisClos <> ChMois) then
             ExecuteSQL ('UPDATE ANNUAIRE SET ANN_MOISCLOTURE="'+ChMois+'" WHERE ANN_GUIDPER="'+guidperdos+'"');
     end;
   end;

   // Lancement effectif de la fiche
   if Inside = nil then   //+LMO
     AGLLanceFiche(Nature, Fiche, Range, Lequel, Argument)
   else
   begin
     AGLLanceFicheInside(Nature, Fiche, Range, Lequel, Argument, inside);
   end ;                  //-LMO
end;


procedure UpdateParamSocDossier(NomParam, Valeur: String);
// maj d'un paramsoc dans la base du dossier en cours
// #### ne marche que sur la base commune
begin
  if (V_PGI.RunFromLanceur) or ( (V_PGI.Driver<>dbMSSQL) and (V_PGI.Driver<>dbMSSQL2005) ) then exit;

  if V_PGI.DefaultSectionName<>'' then
    if V_PGI.DbName<>V_PGI.DefaultSectionDBName then exit;

  if (VH_Doss.GuidPer='') or (VH_Doss.NoDossier='') then exit;

  if Not DBExists ('DB' + VH_DOSS.NoDossier) then exit; 

  // Paramètres société
  ExecuteSQL('UPDATE '+VH_Doss.DBSocName+'.dbo.PARAMSOC SET SOC_DATA="' + Valeur + '" WHERE SOC_NOM="'+NomParam+'"') ;
end;


procedure AfficheLibTablesLibres(LaTof: TOF);
// affiche les libellés des zones libres (onglet Divers de la fiche annuaire)
var TOBCaption : TOB;
    tmp : String;

  procedure TraiteZoneLibre(etiq1, etiq2, zone1, zone2: String; txt: String);
  begin
    With LaTof do
      begin
      if txt<>'' then
        // renseigne le nom identifié de la zone libre
        SetControlCaption(etiq1, txt)
      else
        // ou désactive les zones libres non utilisées
        begin
        SetControlVisible(etiq1, False);
        if etiq2<>'' then SetControlVisible(etiq2, False);
        SetControlVisible(zone1, False);
        if zone2<>'' then SetControlVisible(zone2, False);
        end;
      end;
  end;

begin
// rq : table ANNUPARAM aurait pu être des paramsoc car un seul enreg
  TOBCaption := TOB.Create('ANNUPARAM',nil,-1);
  TOBCaption.SelectDB('',nil);
  TraiteZoneLibre('lblCharLibre1', '', 'ANB_CHARLIBRE1', '', TOBCaption.GetValue('ANP_TEXTELIBRE1'));
  TraiteZoneLibre('lblCharLibre2', '', 'ANB_CHARLIBRE2', '', TOBCaption.GetValue('ANP_TEXTELIBRE2'));
  TraiteZoneLibre('lblCharLibre3', '', 'ANB_CHARLIBRE3', '', TOBCaption.GetValue('ANP_TEXTELIBRE3'));
  TraiteZoneLibre('lblDateLibre1', 'lblDateLibreA1', 'ANB_DATELIBRE1', 'ANB_DATELIBRE1_', TOBCaption.GetValue('ANP_DATELIBRE1'));
  TraiteZoneLibre('lblDateLibre2', 'lblDateLibreA2', 'ANB_DATELIBRE2', 'ANB_DATELIBRE2_', TOBCaption.GetValue('ANP_DATELIBRE2'));
  TraiteZoneLibre('lblDateLibre3', 'lblDateLibreA3', 'ANB_DATELIBRE3', 'ANB_DATELIBRE3_', TOBCaption.GetValue('ANP_DATELIBRE3'));
  TraiteZoneLibre('lblValLibre1', 'lblValLibreA1', 'ANB_VALLIBRE1', 'ANB_VALLIBRE1_', TOBCaption.GetValue('ANP_MONTANTLIBRE1'));
  TraiteZoneLibre('lblValLibre2', 'lblValLibreA2', 'ANB_VALLIBRE2', 'ANB_VALLIBRE2_', TOBCaption.GetValue('ANP_MONTANTLIBRE2'));
  TraiteZoneLibre('lblValLibre3', 'lblValLibreA3', 'ANB_VALLIBRE3', 'ANB_VALLIBRE3_', TOBCaption.GetValue('ANP_MONTANTLIBRE3'));
  TraiteZoneLibre('ANB_BOOLLIBRE1', '', 'ANB_BOOLLIBRE1', '', TOBCaption.GetValue('ANP_COCHELIBRE1'));
  TraiteZoneLibre('ANB_BOOLLIBRE2', '', 'ANB_BOOLLIBRE2', '', TOBCaption.GetValue('ANP_COCHELIBRE2'));
  TraiteZoneLibre('ANB_BOOLLIBRE3', '', 'ANB_BOOLLIBRE3', '', TOBCaption.GetValue('ANP_COCHELIBRE3'));
  tmp := RechDom('YYANNUPARAM','AN1',false);
  if tmp='Tablette n°1 Annuaire' then tmp := '';
  TraiteZoneLibre('lblChoixLibre1', '', 'ANB_CHOIXLIBRE1', '', tmp);
  tmp := RechDom('YYANNUPARAM','AN2',false);
  if tmp='Tablette n°2 Annuaire' then tmp := '';
  TraiteZoneLibre('lblChoixLibre2', '', 'ANB_CHOIXLIBRE2', '', tmp);
  tmp := RechDom('YYANNUPARAM','AN3',false);
  if tmp='Tablette n°3 Annuaire' then tmp := '';
  TraiteZoneLibre('lblChoixLibre3', '', 'ANB_CHOIXLIBRE3', '', tmp);

  // $$$ JP 04/04/06
  tmp := RechDom ('YYZONELIBREGED', 'ZL1', FALSE);
  TraiteZoneLibre ('TYDO_LIBREGED1', '', 'YDO_LIBREGED1', '', tmp);
  tmp := RechDom ('YYZONELIBREGED', 'ZL2', FALSE);
  TraiteZoneLibre ('TYDO_LIBREGED2', '', 'YDO_LIBREGED2', '', tmp);
  tmp := RechDom ('YYZONELIBREGED', 'ZL3', FALSE);
  TraiteZoneLibre ('TYDO_LIBREGED3', '', 'YDO_LIBREGED3', '', tmp);
  tmp := RechDom ('YYZONELIBREGED', 'ZL4', FALSE);
  TraiteZoneLibre ('TYDO_LIBREGED4', '', 'YDO_LIBREGED4', '', tmp);
  tmp := RechDom ('YYZONELIBREGED', 'ZL5', FALSE);
  TraiteZoneLibre ('TYDO_LIBREGED5', '', 'YDO_LIBREGED5', '', tmp);

  TobCaption.Free;
end;


function IsTabEmpty (Tab:TTabSheet):boolean;
var
   i    :integer;
begin
    Result := FALSE;
    If Tab=Nil then
        begin
        Result := TRUE;
        exit; //mcd 13/12/2005 .. plante si le controle n'existe pas
        end;
    for i := 0 to Tab.ControlCount -1 do
    begin
        if Tab.Controls [i].Visible = TRUE then
            exit;
    end;

    // Si on arrive ici, il n'y a aucun controle visible dans l'onglet
    Result := TRUE;
end;


function  GuidPerToCodeTiers(sGuidPer: String): String;
var Q: TQuery;
begin
  Result := '';
  Q := OpenSQL('select ANN_TIERS from ANNUAIRE where ANN_GUIDPER="'+sGuidPer+'"', True);
  if Not Q.Eof then Result := Q.Fields[0].AsString;
  Ferme(Q);
end;


function IsPersonneCabinet(sGuidPer_p: String): Boolean;
var Q: TQuery;
begin
  Result := False;
  Q := OpenSQL('select DOS_CABINET from DOSSIER where DOS_GUIDPER = "' + sGuidPer_p + '"', True);
  if Not Q.Eof then
    Result := (Q.FindField('DOS_CABINET').AsString='X');
  Ferme(Q);
end;

{***********A.G.L.***********************************************
Auteur  ...... : MD
Créé le ...... : 21/01/2004
Modifié le ... :   /  /
Description .. : Récupère les infos du principal dirigeant
Suite ........ : du dossier.
Mots clefs ... :
*****************************************************************}
function GetDirigeant(sGuidPer, sFormeJur : String; var sFctGrt, sFctLib, sNomDirig: String): String;
var Q : TQuery;
begin
  Result := '';

  // cherche la fonction du principal dirigeant (JFT_PRIORITE=1)
  // (la plupart du temps : Gérant (GRT), PDG (PCA), ou président du directoire (PDI)
  // ok à partir socref > 591
  sFctGrt := ''; sFctLib := '';
  Q := OpenSQL('select JFT_FONCTION, JTF_FONCTABREGE from JUFONCTION, JUTYPEFONCT'
   +' where JFT_FONCTION=JTF_FONCTION '
   +' and JFT_FORME="'+sFormeJur+'" and JTF_DIRIGEANT="X" and JFT_PRIORITE=1', True);
  if Not Q.Eof then
    begin
    sFctGrt := Q.FindField('JFT_FONCTION').AsString;
    sFctLib := Q.FindField('JTF_FONCTABREGE').AsString;
    end;
  Ferme(Q);

  // identification du Gérant ou Président
  sNomDirig := '';
  // #### attention : limité à la fonction du principal dirigeant
  // mais certaines formes acceptent plusieurs dirigeant (4 pour SA à directoire)
  if sFctGrt<>'' then
    begin
    Q := OpenSQL('select ANL_GUIDPER, ANL_NOMPER from ANNULIEN where ANL_FONCTION="'+sFctGrt
     +'" AND ANL_GUIDPERDOS="'+sGuidPer+'" AND ANL_TYPEDOS="STE" AND ANL_NOORDRE=1',True);
    if not Q.Eof then
      begin
      Result := Q.FindField('ANL_GUIDPER').AsString;
      sNomDirig := Q.FindField('ANL_NOMPER').AsString;
      end;
    Ferme(Q);

    // nom exact ?
    if Result<>'' then
      begin
      Q := OpenSQL('select ANN_PPPM, ANN_CVA, ANN_NOM1, ANN_NOM2 from ANNUAIRE where ANN_GUIDPER = "'
       + Result + '"', True) ;
      if Not Q.Eof then
        begin
        // Personne physique : récupère la civilité
        if Q.FindField('ANN_PPPM').AsString='PP' then sNomDirig := Q.FindField('ANN_CVA').AsString;
        // Retourne Civilité + Prénom + Nom
        sNomDirig := Trim(sNomDirig + ' ' + Q.FindField('ANN_NOM2').AsString + ' ' + Q.FindField('ANN_NOM1').AsString);
        end;
      Ferme(Q);
      end;
    end;
end;

procedure DPLibelleLibres;
var
   iTable, iChamp    :integer;
   strLibelle        :string;
   TOBLibelle        :TOB;
  Mcd : IMCDServiceCOM;
  Table     : ITableCOM ;
Field     : IFieldCOMEx ;
begin
  MCD := TMCD.GetMcd;
  if not mcd.loaded then mcd.WaitLoaded();

     // TOB des libellés issus de la table ANNUPARAM (sauf pour les choix libres, dans tablette)
     TOBLibelle := TOB.Create ('ANNUPARAM', nil, -1);
     try
        // Chargement des libellés en mémoire
        TOBLibelle.SelectDB ('', nil);

        // Màj explicite des champs libres
				field := mcd.getField('ANB_VALLIBRE1') as IfieldComEx;
        if Assigned(field) then
        begin
          strLibelle := Trim (TOBLibelle.GetString ('ANP_MONTANTLIBRE1'));
          if strLibelle <> '' then field.Libelle := strLibelle;
        end;
        //
				field := mcd.getField('ANB_VALLIBRE2') as IfieldComEx;
        if Assigned(field) then
        begin
          strLibelle := Trim (TOBLibelle.GetString ('ANP_MONTANTLIBRE2'));
          if strLibelle <> '' then field.Libelle := strLibelle;
        end;
        //
				field := mcd.getField('ANB_VALLIBRE3') as IfieldComEx;
        if Assigned(field) then
        begin
          strLibelle := Trim (TOBLibelle.GetString ('ANP_MONTANTLIBRE3'));
          if strLibelle <> '' then field.Libelle := strLibelle;
        end;
        //
				field := mcd.getField('ANB_CHARLIBRE1') as IfieldComEx;
        if Assigned(field) then
        begin
          strLibelle := Trim (TOBLibelle.GetString ('ANP_TEXTELIBRE1'));
          if strLibelle <> '' then field.Libelle := strLibelle;
        end;
        //
				field := mcd.getField('ANB_CHARLIBRE2') as IfieldComEx;
        if Assigned(field) then
        begin
          strLibelle := Trim (TOBLibelle.GetString ('ANP_TEXTELIBRE2'));
          if strLibelle <> '' then field.Libelle := strLibelle;
        end;
        //
				field := mcd.getField('ANB_CHARLIBRE3') as IfieldComEx;
        if Assigned(field) then
        begin
          strLibelle := Trim (TOBLibelle.GetString ('ANP_TEXTELIBRE3'));
          if strLibelle <> '' then field.Libelle := strLibelle;
        end;
        //
				field := mcd.getField('ANB_DATELIBRE1') as IfieldComEx;
        if Assigned(field) then
        begin
          strLibelle := Trim (TOBLibelle.GetString ('ANP_DATELIBRE1'));
          if strLibelle <> '' then field.Libelle := strLibelle;
        end;
        //
				field := mcd.getField('ANB_DATELIBRE2') as IfieldComEx;
        if Assigned(field) then
        begin
          strLibelle := Trim (TOBLibelle.GetString ('ANP_DATELIBRE2'));
          if strLibelle <> '' then field.Libelle := strLibelle;
        end;
        //
				field := mcd.getField('ANB_DATELIBRE3') as IfieldComEx;
        if Assigned(field) then
        begin
          strLibelle := Trim (TOBLibelle.GetString ('ANP_DATELIBRE3'));
          if strLibelle <> '' then field.Libelle := strLibelle;
        end;
        //
				field := mcd.getField('ANB_BOOLLIBRE1') as IfieldComEx;
        if Assigned(field) then
        begin
          strLibelle := Trim (TOBLibelle.GetString ('ANP_COCHELIBRE1'));
          if strLibelle <> '' then field.Libelle := strLibelle;
        end;
        //
				field := mcd.getField('ANB_BOOLLIBRE2') as IfieldComEx;
        if Assigned(field) then
        begin
          strLibelle := Trim (TOBLibelle.GetString ('ANP_COCHELIBRE2'));
          if strLibelle <> '' then field.Libelle := strLibelle;
        end;
        //
				field := mcd.getField('ANB_BOOLLIBRE3') as IfieldComEx;
        if Assigned(field) then
        begin
          strLibelle := Trim (TOBLibelle.GetString ('ANP_COCHELIBRE3'));
          if strLibelle <> '' then field.Libelle := strLibelle;
        end;
     finally
            TOBLibelle.Free;
     end;

     // Pour les choix, le libellé est connu dans la tablette YYANNUPARAM
     field := mcd.getField('ANB_CHOIXLIBRE1') as IfieldComEx;
     if Assigned(field) then
     begin
       strLibelle := Trim (RechDom ('YYANNUPARAM', 'AN1', FALSE));
       if strLibelle <> '' then field.Libelle := strLibelle;
     end;
     //
     field := mcd.getField('ANB_CHOIXLIBRE2') as IfieldComEx;
     if Assigned(field) then
     begin
       strLibelle := Trim (RechDom ('YYANNUPARAM', 'AN2', FALSE));
       if strLibelle <> '' then field.Libelle := strLibelle;
     end;
     //
     field := mcd.getField('ANB_CHOIXLIBRE3') as IfieldComEx;
     if Assigned(field) then
     begin
       strLibelle := Trim (RechDom ('YYANNUPARAM', 'AN3', FALSE));
       if strLibelle <> '' then field.Libelle := strLibelle;
     end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 07/06/2006
Modifié le ... : 25/07/2006
Description .. : Permet de se replacer sur l'enregistrement
Suite ........ : qu'on vient de valider quand on est en inside
Suite ........ : (évite la perte de pointeur si double validation)
Mots clefs ... :
*****************************************************************}
procedure ReloadTomInsideAfterInsert(MyFiche_p : TFFiche; DS_P : TDataSet;
                                     aosKeys_p : array of string;
                                     aovCode_p : array of variant);
{$IFNDEF EAGL}
var
   vCode_l : variant;
   sKeys_l : string;
   iInd_l : integer;
{$ENDIF}
begin
   if (MyFiche_p.TypeAction in [taCreat..taCreatOne]) then
   begin
      MyFiche_p.TypeAction := taModif;
      MyFiche_p.OM.RefreshDB;
//      MyFiche_p.ChargeEnreg;
      {$IFDEF EAGL}
      MyFiche_p.ReloadDb;
      {$ELSE}
      if Length(aovCode_p) > 1 then
      begin
         sKeys_l := '';
         vCode_l := VarArrayCreate([0, 0], varVariant);
         for iInd_l := 0 to Length(aovCode_p) - 1 do
         begin
            if sKeys_l <> '' then
               sKeys_l := sKeys_l + ';';
            sKeys_l := sKeys_l + aosKeys_p[iInd_l];
            VarArrayRedim(vCode_l, iInd_l + 1);
            vCode_l[iInd_l] := aovCode_p[iInd_l];
         end;
      end
      else
      begin
         sKeys_l := aosKeys_p[0];
         vCode_l := aovCode_p[0];
      end;

      MyFiche_p.Modifier := false;
      DS_p.Locate(sKeys_l, vCode_l, []);
      {$ENDIF EAGL}
   end;
end;

// $$$ JP 17/10/06: identifiant base commune (GUID), créer si pas encore défini
function dpGetBaseGuid:string;
begin
     Result := GetParamSocSecur ('SO_GUIDDOSSIER', '');
     if Result = '' then
     begin
          SetParamSoc ('SO_GUIDDOSSIER', AGLGetGuid);
          Result := GetParamSocSecur ('SO_GUIDDOSSIER', '');
     end;
end;

// $$$ JP 26/10/06
function BlobToString (Texte:string):string;
var
   Lignes    :HTStrings;
   RichEdit  :TRichEdit;
   Panel     :TPanel;
begin
    Panel := TPanel.Create (nil);
    Panel.Visible := False;
    Panel.ParentWindow := GetDesktopWindow;
    RichEdit := TRichEdit.Create(Panel);
    RichEdit.Parent := Panel;
    Lignes := HTStringList.Create;
    Lignes.Text := Texte;
    StringsToRich (RichEdit, Lignes);
    Result := Trim (RichEdit.Text);
    Lignes.Free;
    RichEdit.Free;
    Panel.Free;

    // On remplace les saut de lignes et tabulations pour que ça passe dans le fichier d'échange (et pas de saut de ligne en fin de texte)
    if Result <> '' then
    begin
         while (Result [Length (Result)] = #10) or (Result [Length (Result)] = #13) do
               Delete (Result, Length (Result), 1);
         Result := StringReplace (Result, #9, ' ', [rfReplaceAll, rfIgnoreCase]);
         Result := StringReplace (Result, #10, '~~', [rfReplaceAll, rfIgnoreCase]);
         Result := StringReplace (Result, #13, '', [rfReplaceAll, rfIgnoreCase]);
    end;
end;


function FaitLeSelect (Ecran:TForm): string ; //LM20070404
//tokenise (virgule) toutes les zones référencée MCD d'une forme
var i : integer ;
    cp : TComponent ;
    st : string ;
begin

  for i:=0 to Ecran.componentCount-1 do
  begin
    cp:=Ecran.components[i] ;
    try
      //pgibox(TControl(cp).Name + ' = ' + intTostr(ChampToNum(TControl(cp).Name)));
      if (TControl(cp).Name<>'') and (ChampToNum(TControl(cp).Name)>-1) then
        st := st + TControl(cp).Name +', ' ;
    except
      on e : Exception do if v_pgi.sav then pgiInfo(e.message)
    end ;
  end ;

  if length(st)>0 then st:= copy(st,1, length(st)-2) ;
  result:=st ;
end ;

procedure GroupBoxEnabled (frm:TForm; gBox:TControl; Sauf : string; ok:Boolean) ;//LM20070511
var i : integer ;
    cp : TComponent ;
begin
  if gBox=nil then exit ;
  sauf:=upperCase(sauf) ;
  for i:=0 to frm.ComponentCount-1 do
  begin
    cp := frm.Components[i] ;
    if (upperCase(cp.name) = sauf)
     or (not cp.InheritsFrom(TControl)) //dispo propriété .parent...
      or (not IsParent(TWincontrol(gBox), TControl(cp))) then continue ;
    if isPropertyExists(cp, 'enabled') then TControl(cp).Enabled := ok ;
  end ;
end ;


function ControleCleFRP( sNoAdmin_p, sNoAdhesion_p, sNoDhCompl_p : string ) : boolean;
var
   sCodeFRP_l : string;
begin
   result := true;
   if ( sNoDhCompl_p = '' ) then
      exit;

   sCodeFRP_l := sNoAdmin_p + sNoAdhesion_p + sNoDhCompl_p;
   result :=  CtrlCodeFRP( false, sCodeFRP_l );
end;

end.

