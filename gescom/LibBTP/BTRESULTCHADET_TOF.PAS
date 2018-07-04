{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 10/05/2010
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BTRESULTCHADET ()
Mots clefs ... : TOF;BTRESULTCHADET
*****************************************************************}
Unit BTRESULTCHADET_TOF ;

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
     uTob,
     Stat,
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     HTB97,
     UTOF,
     variants ;

Type
  TOF_BTRESULTCHADET = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  private
    DateMvtDeb,DateMvtFin : TdateTime;
    Codechantier, NomChoixCod, NomChoixCod2 : string;
    TypeRessource : String;
    TOBBTB : TOB;

    procedure ChargelesInfosConsos;
    procedure DispatchMontants(TOBTMP,TOBECLATE : TOB);
    procedure SetChampsCommun (TOBTMP,OneTOB : TOB);
    procedure AlimenteInfoProv (TOBL,TOBF : TOB);

    procedure AjouteChampSup(var TOBTMP: TOB);
    procedure ChargelesInfosprevu(NaturePiece: string ; WherePiece: string = '');
    procedure DefiniPrevuDetail(TOBPiece: TOB; NaturePiece: string);
    procedure ChargelesLignesTB(TOBLIgne, TOBPiece: TOB);
    procedure ChargelesOuvragesTB(TOBOuvrage, TOBPiece: TOB);
    procedure ConstitueOuvragesTB(TOBPiece, TOBOuvrage,
      TOBLocOuvrage: TOB);
    procedure TraiteLaPieceTB(TOBBTB, TOBPiece, TOBOuvrage: TOB);
    procedure TraiteOuvrageTB(TOBBTB, TOBL, TOBOuvrage: TOB);
    procedure TraiteLigneTB(TOBBTB, TOBL: TOB);
    procedure TraiteDetailOuvrageTB(TOBBTB, TOBOUV: TOB; Qte,QteDuDetail: double);
    procedure TraiteLigneDetailOuvrageTB(TOBBTB, TOBL: TOB; Qte,QteDuDetail: double);
    procedure CumuleprevufactureAutre(TOBTMP: TOB; NaturePiece: string;
      MontantPa, MontantPr, MontantPV, quantite: double);
    procedure CumuleprevufactureFourniture(TOBTMP: TOB;
      NaturePiece: string; MontantPa, MontantPr, MontantPV,TpsPrevu: double);
    procedure CumuleprevufactureInterimaire(TOBTMP: TOB;
      NaturePiece: string; MontantPa, MontantPr, MontantPV,
      TpsPrevu: double);
    procedure CumuleprevufactureLocation(TOBTMP: TOB; NaturePiece: string;
      MontantPa, MontantPr, MontantPV,TpsPRevu: double);
    procedure CumuleprevufactureMateriel(TOBTMP: TOB; NaturePiece: string;
      MontantPa, MontantPr, MontantPV,TpsPrevu: double);
    procedure CumuleprevufactureOutillage(TOBTMP: TOB; NaturePiece: string;
      MontantPa, MontantPr, MontantPV,TpsPrevu: double);
    procedure CumuleprevufactureSalarie(TOBTMP: TOB; NaturePiece: string;
      MontantPa, MontantPr, MontantPV, TpsPrevu: double);
    procedure CumuleprevufactureSousTraitance(TOBTMP: TOB;
      NaturePiece: string; MontantPa, MontantPr, MontantPV,
      TpsPrevu: double);
    procedure ChargelesInfosprevuAvanc;
    procedure ChargeLeFactured;
    procedure RecupBasePrincipale;
    procedure ChargelesInfosEngage;
    procedure DateTransfoOK (TOBEclate : TOB; var Montant : Double);
    procedure RechTiersFou(CodeFrs: String; var LibTiers: String);
    Procedure ChargementEngageFournisseur(TOBTMP,TobProv : TOB);

  end ;

Implementation
  uses  Paramsoc,
        TiersUtil;

procedure TOF_BTRESULTCHADET.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BTRESULTCHADET.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BTRESULTCHADET.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_BTRESULTCHADET.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_BTRESULTCHADET.OnArgument (S : String ) ;
begin
  Inherited ;
  if LaTOB <> nil then
  begin
  	CodeChantier := LaTOB.getValue('CHANTIER');
  	DateMvtDeb := LaTOB.getValue('DATEMVTDEB');
  	DateMvtFin := LaTOB.getValue('DATEMVTFIN');
    Ecran.Caption := 'TABLEAU DE BORD DETAILLE DU CHANTIER : ' +
                     LaTOB.getValue('CHANTIER1')+' '+ LaTOB.getValue('CHANTIER2')+' '+ LaTOB.getValue('CHANTIER3')+
                     ' '+LaTOB.getValue('LIBELLECHANTIER');
  end;
  //
  if CodeChantier <> '' then
  begin
    RecupBasePrincipale;
    TOBBTB := TOB.create('Ma Tob', nil, -1);
    ChargelesInfosConsos;
    ChargelesInfosprevu('DBT');
    ChargelesInfosprevu('BCE');
    ChargelesInfosprevuAvanc;
    ChargeLeFactured;
    ChargelesInfosEngage;
    if TOBBTB <> nil then
    begin
      TFStat(Ecran).LaTOB := TOBBTB;
    end;
  end;

  TypeRessource := GetParamSocSecur('SO_BTNATPRESTATION', '');
  if TypeRessource = '' then
  begin
    pgiinfo('Parametre nature de prestation non renseigné. Traitement impossible');
    ecran.close;
  end;

end ;

procedure TOF_BTRESULTCHADET.RecupBasePrincipale;
var Q: Tquery;
begin
  NomChoixcod := 'CHOIXCOD';
  Q := OpenSQL('SELECT DS_NOMBASE FROM DESHARE WHERE DS_NOMTABLE="GCFAMILLENIV1"', True);
  if not Q.EOF then
  begin
  	NomChoixCod := Q.FindField('DS_NOMBASE').AsString+'.DBO.CHOIXCOD';
  end;
	Ferme (Q);
  NomChoixcod2 := 'CHOIXCOD';
  Q := OpenSQL('SELECT DS_NOMBASE FROM DESHARE WHERE DS_NOMTABLE="AFTTYPERESSOURCE"', True);
  if not Q.EOF then
  begin
  	NomChoixCod2 := Q.FindField('DS_NOMBASE').AsString+'.DBO.CHOIXCOD';
  end;
	Ferme (Q);
end;

procedure TOF_BTRESULTCHADET.OnClose ;
begin
  Inherited ;
  TOBBTB.free;
end ;

procedure TOF_BTRESULTCHADET.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BTRESULTCHADET.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_BTRESULTCHADET.CumuleprevufactureAutre (TOBTMP: TOB; NaturePiece : string; MontantPa,MontantPr,MontantPV,quantite : double);
var LocNaturePiece : string;
begin
  if (Pos(NaturePiece ,'FBT;ABT;FAC;AVC;FBP')>0) Then
  begin
		TOBTMP.PutValue('FACTURE', MontantPV);
  end else
  begin
    if NaturePiece = 'AFF' Then LocNaturePiece := 'PBT' else LocNaturePiece := naturePiece;
    TOBTMP.PutValue('PREVU'+LocNaturePiece+'PA', MontantPA);
    TOBTMP.PutValue('PREVU'+LocNaturePiece+'PR', MontantPR);
    TOBTMP.PutValue('PREVU'+LocNaturePiece+'PV', MontantPV);
		TOBTMP.PutValue('QTEPREVU'+LocNaturePiece,quantite);
  end;
end;

procedure TOF_BTRESULTCHADET.CumuleprevufactureSalarie (TOBTMP : TOB ; NaturePiece : string; MontantPa,MontantPr,MontantPV,TpsPrevu : double);
var LocNaturePiece : string;
begin
//  if (NaturePiece = 'FBT') or (NaturePiece = 'ABT') or (NaturePiece = 'FAC') or (NaturePiece = 'AVC') Then
  if (Pos(NaturePiece ,'FBT;ABT;FAC;AVC;FBP')>0) Then
  begin
     TOBTMP.PutValue('FACTURE', MontantPV);
  end else
  begin
     if NaturePiece = 'AFF' Then LocNaturePiece := 'PBT' else LocNaturePiece := naturePiece;
     TOBTMP.PutValue('PREVU'+LocNaturePiece+'PA',MontantPA);
     TOBTMP.PutValue('PREVU'+LocNaturePiece+'PR',MontantPR);
     TOBTMP.PutValue('PREVU'+LocNaturePiece+'PV',MontantPV);
     TOBTMP.PutValue('QTEPREVU'+LocNaturePiece,TpsPRevu);
  end;
end;

procedure TOF_BTRESULTCHADET.CumuleprevufactureInterimaire (TOBTMP : TOB ; NaturePiece : string; MontantPa,MontantPr,MontantPV,TpsPrevu : double);
var LocNaturePiece : string;
begin
//  if (NaturePiece = 'FBT') or (NaturePiece = 'ABT') or (NaturePiece = 'FAC') or (NaturePiece = 'AVC') Then
  if (Pos(NaturePiece ,'FBT;ABT;FAC;AVC;FBP')>0) Then
  begin
     TOBTMP.PutValue('FACTURE', MontantPV);
  end else
  begin
     if NaturePiece = 'AFF' Then LocNaturePiece := 'PBT' else LocNaturePiece := naturePiece;
     TOBTMP.PutValue('PREVU'+LocNaturePiece+'PA',MontantPA);
     TOBTMP.PutValue('PREVU'+LocNaturePiece+'PR',MontantPR);
     TOBTMP.PutValue('PREVU'+LocNaturePiece+'PV',MontantPV);
     TOBTMP.PutValue('QTEPREVU'+LocNaturePiece,TpsPRevu);
  end;
end;

procedure TOF_BTRESULTCHADET.CumuleprevufactureLocation (TOBTMP : TOB ; NaturePiece : string; MontantPa,MontantPr,MontantPV,TpsPRevu : double);
var LocNaturePiece : string;
begin
//  if (NaturePiece = 'FBT') or (NaturePiece = 'ABT') or (NaturePiece = 'FAC') or (NaturePiece = 'AVC') Then
  if (Pos(NaturePiece ,'FBT;ABT;FAC;AVC;FBP')>0) Then
  begin
     TOBTMP.PutValue('FACTURE', MontantPV);
  end else
  begin
     if NaturePiece = 'AFF' Then LocNaturePiece := 'PBT' else LocNaturePiece := naturePiece;
     TOBTMP.PutValue('PREVU'+LocNaturePiece+'PA',MontantPA);
     TOBTMP.PutValue('PREVU'+LocNaturePiece+'PR',MontantPR);
     TOBTMP.PutValue('PREVU'+LocNaturePiece+'PV',MontantPV);
     TOBTMP.PutValue('QTEPREVU'+LocNaturePiece,TpsPRevu);
  end;
end;

procedure TOF_BTRESULTCHADET.CumuleprevufactureMateriel (TOBTMP : TOB; NaturePiece : string; MontantPa,MontantPr,MontantPV,TpsPrevu : double);
var LocNaturePiece : string;
begin
//  if (NaturePiece = 'FBT') or (NaturePiece = 'ABT') or (NaturePiece = 'FAC') or (NaturePiece = 'AVC') Then
  if (Pos(NaturePiece ,'FBT;ABT;FAC;AVC;FBP')>0) Then
  begin
     TOBTMP.PutValue('FACTURE', MontantPV);
  end else
  begin
     if NaturePiece = 'AFF' Then LocNaturePiece := 'PBT' else LocNaturePiece := naturePiece;
     TOBTMP.PutValue('PREVU'+LocNaturePiece+'PA',MontantPA);
     TOBTMP.PutValue('PREVU'+LocNaturePiece+'PR',MontantPR);
     TOBTMP.PutValue('PREVU'+LocNaturePiece+'PV',MontantPV);
     TOBTMP.PutValue('QTEPREVU'+LocNaturePiece,TpsPRevu);
  end;
end;

procedure TOF_BTRESULTCHADET.CumuleprevufactureOutillage (TOBTMP : TOB; NaturePiece : string; MontantPa,MontantPr,MontantPV,TpsPrevu : double);
var LocNaturePiece : string;
begin
//  if (NaturePiece = 'FBT') or (NaturePiece = 'ABT') or (NaturePiece = 'FAC') or (NaturePiece = 'AVC') Then
  if (Pos(NaturePiece ,'FBT;ABT;FAC;AVC;FBP')>0) Then
  begin
     TOBTMP.PutValue('FACTURE', MontantPV);
  end else
  begin
     if NaturePiece = 'AFF' Then LocNaturePiece := 'PBT' else LocNaturePiece := naturePiece;
     TOBTMP.PutValue('PREVU'+LocNaturePiece+'PA',MontantPA);
     TOBTMP.PutValue('PREVU'+LocNaturePiece+'PR',MontantPR);
     TOBTMP.PutValue('PREVU'+LocNaturePiece+'PV',MontantPV);
     TOBTMP.PutValue('QTEPREVU'+LocNaturePiece,TpsPRevu);
  end;
end;

procedure TOF_BTRESULTCHADET.CumuleprevufactureSousTraitance (TOBTMP : TOB; NaturePiece : string ;MontantPa,MontantPr,MontantPV,TpsPrevu : double);
var LocNaturePiece : string;
begin
//  if (NaturePiece = 'FBT') or (NaturePiece = 'ABT') or (NaturePiece = 'FAC') or (NaturePiece = 'AVC') Then
  if (Pos(NaturePiece ,'FBT;ABT;FAC;AVC;FBP')>0) Then
  begin
     TOBTMP.PutValue('FACTURE', MontantPV);
  end else
  begin
     if NaturePiece = 'AFF' Then LocNaturePiece := 'PBT' else LocNaturePiece := naturePiece;
     TOBTMP.PutValue('PREVU'+LocNaturePiece+'PA',MontantPA);
     TOBTMP.PutValue('PREVU'+LocNaturePiece+'PR',MontantPR);
     TOBTMP.PutValue('PREVU'+LocNaturePiece+'PV',MontantPV);
     TOBTMP.PutValue('QTEPREVU'+LocNaturePiece,TpsPrevu);
  end;
end;

procedure TOF_BTRESULTCHADET.CumuleprevufactureFourniture (TOBTMP : TOB; NaturePiece : string ; MontantPa,MontantPr,MontantPV,TpsPrevu : double);
var LocNaturePiece : string;
begin
//  if (NaturePiece = 'FBT') or (NaturePiece = 'ABT') or (NaturePiece = 'FAC') or (NaturePiece = 'AVC') Then
  if (Pos(NaturePiece ,'FBT;ABT;FAC;AVC;FBP')>0) Then
  begin
     TOBTMP.PutValue('FACTURE', MontantPV);
  end else
  begin
     if NaturePiece = 'AFF' Then LocNaturePiece := 'PBT' else LocNaturePiece := naturePiece;
     TOBTMP.PutValue('PREVU'+LocNaturePiece+'PA',MontantPA);
     TOBTMP.PutValue('PREVU'+LocNaturePiece+'PR',MontantPR);
     TOBTMP.PutValue('PREVU'+LocNaturePiece+'PV',MontantPV);
     TOBTMP.PutValue('QTEPREVU'+LocNaturePiece,TpsPrevu);
  end;
end;

procedure TOF_BTRESULTCHADET.AlimenteInfoProv (TOBL,TOBF : TOB);
var prov : string;
begin
	TOBL.putValue('FOURNISSEUR',TOBF.GetValue('T_LIBELLE'));
  TOBL.putValue('NUMERO', TOBF.GetValue('GL_NUMERO'));
  TOBL.putValue('DATEPIECE',DateToStr(TOBF.GetValue('GL_DATEPIECE')));
  TOBL.putValue('REFERENCE',TOBF.GetValue('GP_REFINTERNE'));
  if TOBF.GetValue('GL_NATUREPIECEG') = 'CF' then
  begin
    TOBL.putValue('NATUREPIECE', 'Commande');
    TOBL.putValue('DATELIVRAISON', DateToStr(TOBF.GetValue('GL_DATELIVRAISON')));
  end else if TOBF.GetValue('GL_NATUREPIECEG') = 'BLF' then
  begin
    TOBL.putValue('NATUREPIECE', 'Réception');
  end else if TOBF.GetValue('GL_NATUREPIECEG') = 'FF' then
  begin
    TOBL.putValue('NATUREPIECE', 'Facture');
  end;
end;

procedure TOF_BTRESULTCHADET.ChargelesInfosConsos;
var Req : string;
    OneTOB : TOB;
    indice : integer;
    TOBTMP, TOBProv, TOBF : TOB;
    QQ : TQuery;
    StWhere : String;
begin
	OneTOB := TOB.Create ('UNE TOB INTERMEDIAIRE',nil,-1);
  Req := 'SELECT BCO_CODEARTICLE AS CODEARTICLE,BNP_TYPERESSOURCE AS NATUREPRESTATION,BCO_NATUREPIECEG,BCO_NATUREMOUV AS NATUREMOUV, ' +
        'BCO_SOUCHE, BCO_NUMERO, BCO_INDICEG, BCO_NUMORDRE, BCO_NUMMOUV, BCO_TRANSFORME, BCO_TRAITEVENTE, ' +
        'BCO_RESSOURCE, BCO_QUANTITE, BCO_MONTANTACH AS ACHAT, ' +
        'BCO_MONTANTPR AS REVIENT, BCO_MONTANTHT AS VENTE, ARS_TYPERESSOURCE, ' +
        'GA_NATUREPRES, BCO_LIBELLE AS LIBELLEART, '+
        'GA_FAMILLENIV1 AS FAMILLE1, GA_FAMILLENIV2 AS FAMILLE2, GA_FAMILLENIV3 AS FAMILLE3, ' +
        'GP_VIVANTE, '+
        '(SELECT CC_LIBELLE FROM '+NomChoixCod2+' WHERE CC_TYPE="TRE" AND CC_CODE=BNP_TYPERESSOURCE) AS LIBNATURE, '+
        '(SELECT CC_LIBELLE FROM '+NomChoixCod2+' WHERE CC_TYPE="TRE" AND CC_CODE=ARS_TYPERESSOURCE) AS LIBNATURE1, '+
        '(SELECT CC_LIBELLE FROM '+NomChoixCod+'  WHERE CC_TYPE="FN1" AND CC_CODE=GA_FAMILLENIV1) AS LIBFAMILLE1, '+
        '(SELECT CC_LIBELLE FROM '+NomChoixCod+'  WHERE CC_TYPE="FN2" AND CC_CODE=GA_FAMILLENIV2) AS LIBFAMILLE2, '+
        '(SELECT CC_LIBELLE FROM '+NomChoixCod+'  WHERE CC_TYPE="FN3" AND CC_CODE=GA_FAMILLENIV3) AS LIBFAMILLE3 '+
        'FROM CONSOMMATIONS ' +
        'LEFT JOIN ARTICLE ON GA_ARTICLE=BCO_ARTICLE ' +
        'LEFT JOIN NATUREPREST ON BNP_NATUREPRES=GA_NATUREPRES '+
        'LEFT JOIN RESSOURCE ON ARS_RESSOURCE=BCO_RESSOURCE '+
        'LEFT JOIN PIECE ON GP_NATUREPIECEG=BCO_NATUREPIECEG AND GP_SOUCHE=BCO_SOUCHE AND GP_NUMERO=BCO_NUMERO '+
        'WHERE BCO_AFFAIRE="' + Codechantier  + '" '+
//--- MODIF BRL 3/08 : 'AND BCO_TRANSFORME="-" AND BCO_TRAITEVENTE="-" ' +
        'AND BCO_NATUREPIECEG IN ("","AFS","AF","BFC","LBT","BLF","CF","CFR","LFR","FF") '+
//---
  			'AND BCO_DATEMOUV BETWEEN "' + UsDateTime(DateMvtDeb) +
        '" AND "' + UsDateTime(DateMvtFin) + '" ' +
  			'ORDER BY BCO_AFFAIRE, BCO_CODEARTICLE, BCO_NATUREPIECEG, BCO_NATUREMOUV';
  OneTOB.LoadDetailDBFromSQL ('LE DETAIL DU CHANTIER',Req,false);

  TOBProv := TOB.Create ('LES LIGNES FOUR',nil,-1);
  for Indice := 0 to OneTOB.detail.count -1 do
  begin
(* inutile : BRL 8/12
    // On ne traite pas les lignes correspondant à des commandes transformées
    if (OneTOB.detail[Indice].Getvalue('BCO_NATUREPIECEG') = 'CF') and
       (OneTOB.detail[Indice].Getvalue('BCO_TRANSFORME') = 'X') then Continue;

    // On ne traite pas les lignes correspondant à des réceptions transformées et livrées sur chantier
    if (OneTOB.detail[Indice].Getvalue('BCO_NATUREPIECEG') = 'BLF') and
       (OneTOB.detail[Indice].Getvalue('BCO_TRANSFORME') = 'X') and
       (OneTOB.detail[Indice].Getvalue('BCO_TRAITEVENTE') = 'X') then Continue;
*)
    TOBTMP := TOB.create('UNE LIGNE',TOBBTB,-1);
    AjouteChampSup (TOBTmp);
    //
    if (not varisnull(OneTOB.getValue('ARS_TYPERESSOURCE'))) and (OneTOB.detail[Indice].GetString('ARS_TYPERESSOURCE') <> '') then
    begin
       OneTOB.detail[Indice].putvalue('NATUREPRESTATION',OneTOB.detail[Indice].GetString('ARS_TYPERESSOURCE'));
       OneTOB.detail[Indice].putvalue('LIBNATURE',OneTOB.detail[Indice].GetString('LIBNATURE1'));
    end;
    //
    SetChampsCommun (TOBTMP,OneTOB.detail[Indice]);
    DispatchMontants(TOBTMP,OneTOB.detail[Indice]);

    // Traitement pour les fournitures et prestations externes:
    if (OneTOB.detail[Indice].GetValue('NATUREMOUV')='FOU') or (OneTOB.detail[Indice].GetValue('NATUREMOUV')='EXT') then
    begin

      StWhere := 'WHERE GL_NATUREPIECEG="' + OneTOB.detail[Indice].GetValue('BCO_NATUREPIECEG') +'"'+
                 '  AND GL_SOUCHE="' + OneTOB.detail[Indice].GetValue('BCO_SOUCHE')             +'"'+
                 '  AND GL_NUMERO='  + IntToStr(OneTOB.detail[Indice].GetValue('BCO_NUMERO'))   +
                 '  AND GL_INDICEG=' + IntToStr(OneTOB.detail[Indice].GetValue('BCO_INDICEG'))  +
                 '  AND GL_NUMORDRE='+ IntToStr(OneTOB.detail[Indice].GetValue('BCO_NUMORDRE'));

      //
      TOBProv.ClearDetail;
      //FV1 : 27/06/2014 - FS#911 - MATFOR - problème de montant en résultat de chantier
      //Chargement engagé fournisseur Commandes fournisseurs non livrées
      if (OneTOB.detail[Indice].GetValue('BCO_NATUREPIECEG')='CF')  Or
         (OneTOB.detail[Indice].GetValue('BCO_NATUREPIECEG')='CFR') Or
         (OneTOB.detail[Indice].GetValue('BCO_NATUREPIECEG')='BFA') Then
      begin
        Req := 'SELECT GL_QTERESTE*GL_PUHTNETDEV AS ACHAT FROM LIGNE ';
        Req := Req + Stwhere;
        Req := Req + ' AND GL_TYPELIGNE="ART"';
        Req := Req + ' AND GL_AFFAIRE="' + Codechantier + '" ';
        Req := Req + ' AND GL_DATEPIECE BETWEEN "' + UsDateTime(DateMvtDeb) + '" ';
        Req := Req + ' AND "' + UsDateTime(DateMvtFin) + '" ';
        QQ := OpenSQL(Req, true,-1,'',true);
        if not QQ.EOF then
        begin
          TOBProv.SelectDB('',QQ);
          ChargementEngageFournisseur(TOBTMP,TObProv);
        end;
        ferme(QQ);
      end;

      //
      TOBProv.ClearDetail;

      // provenance achats
      if (OneTOB.detail[Indice].Getvalue('BCO_NATUREPIECEG') = 'CF') or
         (OneTOB.detail[Indice].Getvalue('BCO_NATUREPIECEG') = 'BLF') then
      begin
        //traitement des lignes issues de commandes fournisseurs
        Req := 'SELECT GL_NATUREPIECEG, GL_NUMERO, GL_TIERS,GL_DATEPIECE, ' +
               'GL_DATELIVRAISON,GP_REFINTERNE, T_LIBELLE, GLC_NATURETRAVAIL ' +
               ' FROM LIGNE ' +
               ' LEFT JOIN PIECE ON GP_NATUREPIECEG=GL_NATUREPIECEG AND GP_SOUCHE=GL_SOUCHE AND GP_NUMERO=GL_NUMERO '+
               ' LEFT JOIN TIERS ON T_TIERS=GL_TIERS AND T_NATUREAUXI="FOU" '+
               ' LEFT JOIN LIGNECOMPL ON GLC_NATUREPIECEG=GL_NATUREPIECEG ' +
               '  AND GLC_SOUCHE=GL_SOUCHE AND GLC_NUMERO=GL_NUMERO ' +
               '  AND GLC_INDICEG=GL_INDICEG AND GLC_NUMORDRE=GL_NUMORDRE ';
        Req := Req + stWhere;

        QQ := OpenSql (req,true);
        if not QQ.eof then
        begin
          TOBProv.SelectDB('',QQ) ;
          AlimenteInfoProv (TOBTMP,TOBPROV);
        end;
        ferme (QQ);
      end else if (OneTOB.detail[Indice].Getvalue('BCO_NATUREPIECEG') = 'LBT') then
      begin
        // traitement des lignes de consos issues de réceptions ou factures
        Req := 'SELECT GL_NATUREPIECEG, GL_NUMERO, GL_TIERS,GL_DATEPIECE, ' +
               'GL_DATELIVRAISON, GP_REFINTERNE, T_LIBELLE, GLC_NATURETRAVAIL '+
               'FROM CONSOMMATIONS ' +
               'LEFT JOIN LIGNE ON GL_NATUREPIECEG=BCO_NATUREPIECEG ' +
               ' AND GL_SOUCHE=BCO_SOUCHE AND GL_NUMERO=BCO_NUMERO  ' +
               ' AND GL_INDICEG=BCO_INDICEG '+
               'LEFT JOIN PIECE ON GP_NATUREPIECEG=GL_NATUREPIECEG  ' +
               ' AND GP_SOUCHE=GL_SOUCHE AND GP_NUMERO=GL_NUMERO ' +
               'LEFT JOIN TIERS ON T_TIERS=GL_TIERS AND T_NATUREAUXI="FOU" '+
               'LEFT JOIN LIGNECOMPL ON GLC_NATUREPIECEG=GL_NATUREPIECEG ' +
               ' AND GLC_SOUCHE=GL_SOUCHE AND GLC_NUMERO=GL_NUMERO ' +
               ' AND GLC_INDICEG=GL_INDICEG AND GLC_NUMORDRE=GL_NUMORDRE ' +
               'WHERE BCO_LIENVENTE='+floatToStr(OneTOB.detail[Indice].GetValue('BCO_NUMMOUV'));
        QQ := OpenSql (req,true);
        if not QQ.eof then
        begin
          TOBProv.SelectDB('',QQ) ;
          AlimenteInfoProv (TOBTMP,TOBPROV);
        end;
        ferme (QQ);
      end;
    end;
  end;
  if TOBProv <> nil then FreeAndNil(TOBProv);
  OneTOB.free;
end;

procedure TOF_BTRESULTCHADET.AjouteChampSup(var TOBTMP: TOB);
begin
  TOBTMP.addchampsupvaleur('CODEARTICLE', '');
  TOBTMP.addchampsupvaleur('NATURE','');
  TOBTMP.addchampsupvaleur('FAMILLE1', '');
  TOBTMP.addchampsupvaleur('FAMILLE2', '');
  TOBTMP.addchampsupvaleur('FAMILLE3', '');
  // prévu via le devis
  TOBTMP.addchampsup('PREVUDBTPA', false); TOBTMP.PutValue('PREVUDBTPA', 0.0);
  TOBTMP.addchampsup('PREVUDBTPR', false); TOBTMP.PutValue('PREVUDBTPR', 0.0);
  TOBTMP.addchampsup('PREVUDBTPV', false); TOBTMP.PutValue('PREVUDBTPV', 0.0);
  TOBTMP.addchampsup('QTEPREVUDBT', false); TOBTMP.PutValue('QTEPREVUDBT', 0.0);
  // prévu via la contre étude
  TOBTMP.addchampsup('PREVUBCEPA', false); TOBTMP.PutValue('PREVUBCEPA', 0.0);
  TOBTMP.addchampsup('PREVUBCEPR', false); TOBTMP.PutValue('PREVUBCEPR', 0.0);
  TOBTMP.addchampsup('PREVUBCEPV', false); TOBTMP.PutValue('PREVUBCEPV', 0.0);
  TOBTMP.addchampsup('QTEPREVUBCE', false); TOBTMP.PutValue('QTEPREVUBCE', 0.0);
  // prévu via la prévision de chantier
  TOBTMP.addchampsup('PREVUPBTPA', false); TOBTMP.PutValue('PREVUPBTPA', 0.0);
  TOBTMP.addchampsup('PREVUPBTPR', false); TOBTMP.PutValue('PREVUPBTPR', 0.0);
  TOBTMP.addchampsup('PREVUPBTPV', false); TOBTMP.PutValue('PREVUPBTPV', 0.0);
  TOBTMP.addchampsup('QTEPREVUPBT', false); TOBTMP.PutValue('QTEPREVUPBT', 0.0);
  //
  TOBTMP.addchampsup('REALISEPA', false); TOBTMP.PutValue('REALISEPA', 0.0);
  TOBTMP.addchampsup('REALISEPR', false); TOBTMP.PutValue('REALISEPR', 0.0);
  TOBTMP.addchampsup('REALISEPV', false); TOBTMP.PutValue('REALISEPV', 0.0);
  TOBTMP.addchampsup('QTEREALISE', false); TOBTMP.PutValue('QTEREALISE', 0.0);
  //
  TOBTMP.addchampsup('AVANCEPA', false); TOBTMP.PutValue('AVANCEPA', 0.0);
  TOBTMP.addchampsup('AVANCEPR', false); TOBTMP.PutValue('AVANCEPR', 0.0);
  TOBTMP.addchampsup('AVANCEPV', false); TOBTMP.PutValue('AVANCEPV', 0.0);
  TOBTMP.addchampsup('QTEAVANCE', false); TOBTMP.PutValue('QTEAVANCE', 0.0);
  //
  TOBTMP.addchampsup('FACTURE', false); TOBTMP.PutValue('FACTURE', 0.0);
  //
  TOBTMP.addchampsup('ENGAGEFOU', false); TOBTMP.PutValue('ENGAGEFOU', 0.0);
  TOBTMP.addchampsup('ENGAGECHA', false); TOBTMP.PutValue('ENGAGECHA', 0.0);
  TOBTMP.addchampsup('FACTNONPAR', false); TOBTMP.PutValue('FACTNONPAR', 0.0);
  //
  TOBTMP.addchampsup('COUTANNEXE', false); TOBTMP.PutValue('COUTANNEXE', 0.0);
  TOBTMP.addchampsup('RECETTEANNEXE', false); TOBTMP.PutValue('RECETTEANNEXE', 0.0);
  //  champs pour infos provenance achats
  TOBTMP.addchampsupvaleur('FOURNISSEUR', '');
  TOBTMP.addchampsupvaleur('NOM FOURNISSEUR', '');
  TOBTMP.addchampsupvaleur('NATUREPIECE', '');
  TOBTMP.addchampsupvaleur('NUMERO', 0);
  TOBTMP.addchampsupvaleur('DATEPIECE', '');
  TOBTMP.addchampsupvaleur('REFERENCE', '');
  TOBTMP.addchampsupvaleur('DATELIVRAISON', '');
end;

procedure TOF_BTRESULTCHADET.ChargelesLignesTB (TOBLIgne,TOBPiece : TOB);
var Req : string;
begin
	Req := 'SELECT BNP_TYPERESSOURCE AS NATUREPRESTATION,GL_NATUREPIECEG,GL_TYPEARTICLE,GL_TYPENOMENC,GL_CODEARTICLE AS CODEARTICLE,GL_LIBELLE AS LIBELLEART,GL_NUMLIGNE,GL_NUMORDRE,'+
  			 'GL_INDICENOMEN,GL_QTEFACT,GL_DPA,GL_DPR,GL_TOTALHTDEV,GL_PUHTDEV,GL_FOURNISSEUR,GLC_NATURETRAVAIL,' +
         'GA_NATUREPRES ,'+
         'GA_FAMILLENIV1 AS FAMILLE1, GA_FAMILLENIV2 AS FAMILLE2, GA_FAMILLENIV3 AS FAMILLE3, ' +
         '(SELECT CC_LIBELLE FROM '+NomChoixCod2+' WHERE CC_TYPE="TRE" AND CC_CODE=BNP_TYPERESSOURCE) AS LIBNATURE, '+
         '(SELECT CC_LIBELLE FROM '+NomChoixCod+'  WHERE CC_TYPE="FN1" AND CC_CODE=GA_FAMILLENIV1) AS LIBFAMILLE1, '+
         '(SELECT CC_LIBELLE FROM '+NomChoixCod+'  WHERE CC_TYPE="FN2" AND CC_CODE=GA_FAMILLENIV2) AS LIBFAMILLE2, '+
         '(SELECT CC_LIBELLE FROM '+NomChoixCod+'  WHERE CC_TYPE="FN3" AND CC_CODE=GA_FAMILLENIV3) AS LIBFAMILLE3 '+
         'FROM LIGNE '+
  			 'LEFT JOIN ARTICLE ON GA_ARTICLE=GL_ARTICLE '+
         'LEFT JOIN NATUREPREST ON BNP_NATUREPRES=GA_NATUREPRES '+
         'LEFT JOIN LIGNECOMPL ON GLC_NATUREPIECEG=GL_NATUREPIECEG AND ' +
         'GLC_SOUCHE=GL_SOUCHE AND GLC_NUMERO=GL_NUMERO AND ' +
         'GLC_INDICEG=GL_INDICEG AND GLC_NUMORDRE=GL_NUMORDRE ' +
         'WHERE '+
         'GL_TYPELIGNE="ART" ' + // on ne prend que les lignes de type article (les commentaires...pfff)
         'AND GL_NATUREPIECEG = "'+TOBPiece.getValue('GP_NATUREPIECEG')+ '" '+
         'AND GL_SOUCHE="' + TOBPiece.GetValue('GP_SOUCHE') + '" ' +
         'AND GL_NUMERO=' + IntToStr(TOBPiece.GetValue('GP_NUMERO')) + ' '+
         'AND GL_INDICEG=' + IntToStr(TOBPiece.GetValue('GP_INDICEG')) + ' '+
         'ORDER BY GL_NUMLIGNE';
  TOBLigne.LoadDetailDBFromSQL  ('LIGNE',Req,false);
end;

procedure TOF_BTRESULTCHADET.ConstitueOuvragesTB (TOBPiece,TOBOuvrage,TOBLocOuvrage : TOB);
var Lig : integer;
		indice : integer;
    TOBLig ,TOBNOuv, TOBPere, TOBnewDet,TOBL : TOB;
    LigneN1,LigneN2,LigneN3,LigneN4,LigneN5 : integer;
begin
	// Initialisation
	Lig := 0;
  TOBNouv := nil;
  //
	for Indice := 0 TO TOBLocOuvrage.detail.count -1 do
  begin
    TOBLig := TOBLocOuvrage.detail[Indice];
    if TOBLIg.getValue('BLO_NUMLIGNE') <> Lig then
    begin
    	// rupture sur N° de ligne --> donc nouvel ouvrage
      TOBNOuv := TOB.create ('NEW OUV',TOBOuvrage,-1);
      Lig := TOBLIg.getValue('BLO_NUMLIGNE');
      TOBL := TOBPiece.findFirst(['GL_NUMLIGNE'],[Lig],true);
      if TOBL<> nil then TOBL.putValue('GL_INDICENOMEN',TOBOuvrage.detail.count);
    end;
    LigneN1 := TOBLig.GetValue('BLO_N1');
    LigneN2 := TOBLig.GetValue('BLO_N2');
    LigneN3 := TOBLig.GetValue('BLO_N3');
    LigneN4 := TOBLig.GetValue('BLO_N4');
    LigneN5 := TOBLig.GetValue('BLO_N5');

    if LigneN5 > 0 then
    begin
			TOBPere:=TOBNOuv.FindFirst(['BLO_NUMLIGNE','BLO_N1','BLO_N2','BLO_N3','BLO_N4','BLO_N5'],[Lig,LigneN1,LigneN2,LigneN3,LigneN4,0],True) ;
    end else
    if LigneN4 > 0 then
    begin
    	TOBPere:=TOBNOuv.FindFirst(['BLO_NUMLIGNE','BLO_N1','BLO_N2','BLO_N3','BLO_N4','BLO_N5'],[Lig,LigneN1,LigneN2,LigneN3,0,0],True) ;
    end else
    if LigneN3 > 0 then
    begin
    	TOBPere:=TOBNOuv.FindFirst(['BLO_NUMLIGNE','BLO_N1','BLO_N2','BLO_N3','BLO_N4','BLO_N5'],[Lig,LigneN1,LigneN2,0,0,0],True) ;
    end else
    if LigneN2 > 0 then
    begin
      TOBPere:=TOBNOuv.FindFirst(['BLO_NUMLIGNE','BLO_N1','BLO_N2','BLO_N3','BLO_N4','BLO_N5'],[Lig,LigneN1,0,0,0,0],True) ;
    end else
    begin
    	TOBPere:=TOBNOuv;
    end;

    if TOBPere<>Nil then
    BEGIN
       TOBNewDet:=TOB.Create('LIGNEOUV',TOBPere,-1) ;
       TOBNewDet.Dupliquer(TOBLig,False,True) ;
    END;
  end;
end;

procedure TOF_BTRESULTCHADET.ChargelesOuvragesTB (TOBOuvrage,TOBPiece : TOB);
var Req : string;
    TOBLocOuvrage : TOB;
begin
	TOBLocOuvrage := TOB.Create ('OUV LU',nil,-1);
	Req := 'SELECT BNP_TYPERESSOURCE AS NATUREPRESTATION,BLO_TYPEARTICLE,BLO_NATUREPIECEG,BLO_CODEARTICLE AS CODEARTICLE,BLO_QTEFACT,BLO_LIBELLE AS LIBELLEART,BLO_DPA,BLO_REMISEPIED,'+
  			 'BLO_DPR,BLO_PUHTDEV,BLO_QTEDUDETAIL,BLO_NUMLIGNE,BLO_N1,BLO_N2,BLO_N3,BLO_N4,BLO_N5, BLO_NATURETRAVAIL, BLO_FOURNISSEUR, ' +
         'GA_NATUREPRES,GA_FAMILLENIV1 AS FAMILLE1, GA_FAMILLENIV2 AS FAMILLE2, GA_FAMILLENIV3 AS FAMILLE3, ' +
         '(SELECT CC_LIBELLE FROM '+NomChoixCod2+' WHERE CC_TYPE="TRE" AND CC_CODE=BNP_TYPERESSOURCE) AS LIBNATURE, '+
         '(SELECT CC_LIBELLE FROM '+NomChoixCod+' WHERE CC_TYPE="FN1" AND CC_CODE=GA_FAMILLENIV1) AS LIBFAMILLE1, '+
         '(SELECT CC_LIBELLE FROM '+NomChoixCod+' WHERE CC_TYPE="FN2" AND CC_CODE=GA_FAMILLENIV2) AS LIBFAMILLE2, '+
         '(SELECT CC_LIBELLE FROM '+NomChoixCod+' WHERE CC_TYPE="FN3" AND CC_CODE=GA_FAMILLENIV3) AS LIBFAMILLE3 '+
         'FROM LIGNEOUV '+
  			 'LEFT JOIN ARTICLE ON GA_ARTICLE=BLO_ARTICLE '+
         'LEFT JOIN NATUREPREST ON BNP_NATUREPRES=GA_NATUREPRES '+
         'WHERE '+
         'AND BLO_NATUREPIECEG = "'+TOBPiece.getValue('GP_NATUREPIECEG')+'" '+
         'AND BLO_SOUCHE="' + TOBPiece.GetValue('GP_SOUCHE') + '" ' +
         'AND BLO_NUMERO=' + IntToStr(TOBPiece.GetValue('GP_NUMERO')) + ' '+
         'AND BLO_INDICEG=' + IntToStr(TOBPiece.GetValue('GP_INDICEG')) + ' '+
         'ORDER BY BLO_NUMLIGNE,BLO_N1,BLO_N2,BLO_N3,BLO_N4,BLO_N5';
  TOBLocOuvrage.LoadDetailDBFromSQL  ('LIGNEOUV',Req,false);
  ConstitueOuvragesTB (TOBPiece,TOBOuvrage,TOBLocOuvrage);
  TOBLocOuvrage.free;
end;

procedure TOF_BTRESULTCHADET.TraiteLigneTB (TOBBTB,TOBL : TOB);
var MontantPa,MontantPr,MontantPv,TpsPrevu : double;
		NaturePiece   : String;
    LibFournisseur: String;
    TOBTMP : TOB;
begin

  //si intervenant sur ligne article pas prise en compte
  if TOBL.GetString('GLC_NATURETRAVAIL') = '001' then exit;


  if (TobL.getString('GLC_NATURETRAVAIL') = '002') Then
  begin
    TOBL.PutValue('NATUREPRESTATION', 'ST');
    TOBL.PutValue('LIBNATURE', rechDom('AFTTYPERESSOURCE','ST',false));
  end;

	TOBTMP := TOB.create('UNE LIGNE',TOBBTB,-1);

	AjouteChampSup (TOBTMP);
	SetChampsCommun (TOBTMP,TOBL);

  if (TobL.getString('GLC_NATURETRAVAIL') = '002') Then
  begin
    TOBTMP.PutValue('FOURNISSEUR',TOBL.getString('GL_FOURNISSEUR'));
    RechTiersFou(TOBL.getString('GL_FOURNISSEUR'),Libfournisseur);
    TOBTMP.PUTVALUE('NOM FOURNISSEUR', Libfournisseur);
  end;
  //
	NaturePiece := TOBL.GetValue('GL_NATUREPIECEG');
  MontantPA := TOBL.GetValue('GL_QTEFACT')*TOBL.GetValue('GL_DPA');
  MontantPR := TOBL.GetValue('GL_QTEFACT')*TOBL.GetValue('GL_DPR');
  MontantPV := TOBL.GetValue('GL_TOTALHTDEV');
  TpsPrevu := TOBL.GetValue('GL_QTEFACT');
  //
  if TOBL.GetValue('GL_TYPEARTICLE')='POU' then
  begin
    // Article de type pourcentage
    MontantPA := 0 ;
    MontantPR := 0;
    MontantPV := (TpsPrevu*TOBL.GetValue('GL_PUHTDEV')/100);
  end;

  if TOBL.GetString('NATUREPRESTATION') = 'SAL' then
  begin
    // prevu salarie
    CumuleprevufactureSalarie (TOBTMP,NaturePiece,MontantPa,MontantPr,MontantPV,TpsPrevu);
  end else if TOBL.GetString('NATUREPRESTATION') = 'AUT' then
  begin
    // prevu autre
    CumuleprevufactureAutre (TOBTMP,NaturePiece,MontantPa,MontantPr,MontantPV,TpsPrevu);
  end else if TOBL.GetString('NATUREPRESTATION') = 'INT' then
  begin
    // prevu interimaire
    CumuleprevufactureInterimaire (TOBTMP,NaturePiece,MontantPa,MontantPr,MontantPV,TpsPrevu);
  end else if TOBL.GetString('NATUREPRESTATION') = 'LOC' then
  begin
    // prevu location
    CumuleprevufactureLocation (TOBTMP,NaturePiece,MontantPa,MontantPr,MontantPV,TpsPrevu);
  end else if TOBL.GetString('NATUREPRESTATION') = 'MAT' then
  begin
    // prevu materiel
    CumuleprevufactureMateriel (TOBTMP,NaturePiece,MontantPa,MontantPr,MontantPV,TpsPrevu);
  end else if TOBL.GetString('NATUREPRESTATION') = 'OUT' then
  begin
    // prevu outillage
    CumuleprevufactureOutillage (TOBTMP,NaturePiece,MontantPa,MontantPr,MontantPV,TpsPrevu);
  end else if TOBL.GetString('NATUREPRESTATION') = 'ST' then
  begin
    // prevu sous traitance
    CumuleprevufactureSousTraitance (TOBTMP,NaturePiece,MontantPa,MontantPr,MontantPV,TpsPrevu);
  end else
  begin
    // prevu fourniture
    CumuleprevufactureFourniture(TOBTMP,NaturePiece,MontantPa,MontantPr,MontantPV,TpsPrevu);
  end;
end;


procedure TOF_BTRESULTCHADET.TraiteLigneDetailOuvrageTB (TOBBTB,TOBL : TOB;Qte,QteDuDetail : double);
var MontantPa,MontantPr,MontantPv,TpsPrevu : double;
		NaturePiece : string;
    Libfournisseur : String;
    TOBTMp : TOB;
begin

	TOBTMP := TOB.create('UNE LIGNE',TOBBTB,-1);
	AjouteChampSup (TOBTMP);
	SetChampsCommun (TOBTMP,TOBL);
  if (TobL.getString('BLO_NATURETRAVAIL') = '002') Then
  begin
    TOBTMP.PutValue('FOURNISSEUR',TOBL.getString('BLO_FOURNISSEUR'));
    RechTiersFou(TOBL.getString('BLO_FOURNISSEUR'),Libfournisseur);
    TOBTMP.PUTVALUE('NOM FOURNISSEUR', Libfournisseur);
  end;

  //
	NaturePiece := TOBL.GetValue('BLO_NATUREPIECEG');
  TpsPrevu := Qte/QteDudetail;
  MontantPA := (TpsPrevu*TOBL.GetValue('BLO_DPA')) ;
  MontantPR := (TpsPrevu*TOBL.GetValue('BLO_DPR'));
  MontantPV := (TpsPrevu*TOBL.GetValue('BLO_PUHTDEV'));
  if TOBL.GetValue('BLO_TYPEARTICLE')='POU' then
  begin
    // Article de type pourcentage
    MontantPA := 0 ;
    MontantPR := 0;
    MontantPV := (TpsPrevu*TOBL.GetValue('BLO_PUHTDEV')/100);
  end;

	if TOBL.GetValue('BLO_REMISEPIED') then
  begin
  	MontantPV := MontantPV - (MontantPV * (TOBL.GetValue('BLO_REMISEPIED')/100.0));
  end;


  if TOBL.GetValue('NATUREPRESTATION') = 'SAL' then
  begin
    // prevu salarie
    CumuleprevufactureSalarie (TOBTMP,NaturePiece,MontantPa,MontantPr,MontantPV,TpsPrevu);
  end else if TOBL.GetValue('NATUREPRESTATION') = 'AUT' then
  begin
    // prevu autre
    CumuleprevufactureAutre (TOBTMP,NaturePiece,MontantPa,MontantPr,MontantPV,TpsPrevu);
  end else if TOBL.GetValue('NATUREPRESTATION') = 'INT' then
  begin
    // prevu interimaire
    CumuleprevufactureInterimaire (TOBTMP,NaturePiece,MontantPa,MontantPr,MontantPV,TpsPrevu);
  end else if TOBL.GetValue('NATUREPRESTATION') = 'LOC' then
  begin
    // prevu location
    CumuleprevufactureLocation (TOBTMP,NaturePiece,MontantPa,MontantPr,MontantPV,TpsPrevu);
  end else if TOBL.GetValue('NATUREPRESTATION') = 'MAT' then
  begin
    // prevu materiel
    CumuleprevufactureMateriel (TOBTMP,NaturePiece,MontantPa,MontantPr,MontantPV,TpsPrevu);
  end else if TOBL.GetValue('NATUREPRESTATION') = 'OUT' then
  begin
    // prevu outillage
    CumuleprevufactureOutillage (TOBTMP,NaturePiece,MontantPa,MontantPr,MontantPV,TpsPrevu);
  end else if TOBL.GetValue('NATUREPRESTATION') = 'ST' then
  begin
    // prevu sous traitance
    CumuleprevufactureSousTraitance (TOBTMP,NaturePiece,MontantPa,MontantPr,MontantPV,TpsPrevu);
  end else
  begin
    // prevu fourniture
    CumuleprevufactureFourniture (TOBTMP,NaturePiece,MontantPa,MontantPr,MontantPV,TpsPrevu);
  end;
end;


procedure TOF_BTRESULTCHADET.TraiteOuvrageTB (TOBBTB,TOBL,TOBOuvrage : TOB);
var IndiceOuv : integer;
		TOBOuv : TOB;
begin
  IndiceOuv := TOBL.GetValue('GL_INDICENOMEN'); if IndiceOuv =0 then exit;
  TOBOuv := TOBOuvrage.detail[IndiceOuv-1];   if TOBOUv = nil then exit;
  TraiteDetailOuvrageTB (TOBBTB,TOBOUV,TOBL.GetValue('GL_QTEFACT'),1);
end;

procedure TOF_BTRESULTCHADET.TraiteDetailOuvrageTB (TOBBTB,TOBOUV : TOB; Qte,QteDuDetail : double);
var QteSui,QteDuDetailSui : double;
		Indice        : integer;
    TOBDet        : TOB;
begin
  for indice := 0 to TOBOUV.detail.count -1 do
  begin
  	TOBDet := TOBOUV.detail[Indice];
    if (Tobdet.getString('BLO_NATURETRAVAIL') = '001') Then continue;
    if (Tobdet.getString('BLO_NATURETRAVAIL') = '002') Then
    begin
      TOBDet.PutValue('NATUREPRESTATION','ST');
      TOBDet.PutValue('LIBNATURE', rechDom('AFTTYPERESSOURCE','ST',false));
    end;
    QteSui := Qte * TOBDet.GetValue('BLO_QTEFACT');
    if TOBDet.GetValue('BLO_QTEDUDETAIL') <> 0 then
    begin
       QteDuDetailSui := QteDudetail * TOBDet.GetValue('BLO_QTEDUDETAIL');
    end else
    begin
       QteDuDetailSui := QteDudetail;
    end;
    if TOBDet.detail.count > 0 then
    begin
      TraiteDetailOuvrageTB (TOBBTB,TOBDet,QteSui,QteDuDetailSui);
    end else
    begin
    	TraiteLigneDetailOuvrageTB (TOBBTB,TOBdet,QteSui,QteDuDetailSui);
    end;
  end;
end;

procedure TOF_BTRESULTCHADET.TraiteLaPieceTB (TOBBTB,TOBPiece,TOBOuvrage : TOB);
var Indice : integer;
		TOBL : TOB;
begin

	for indice := 0 to TOBPiece.detail.count -1 do
  begin
    TOBL := TOBPiece.detail[Indice];
    if (Pos (TOBL.GetValue('GL_TYPEARTICLE'),'OUV;OU1;ARP') > 0) and (TOBL.GetValue('GL_INDICENOMEN')>0) then
    begin
    	TraiteOuvrageTB (TOBBTB,TOBL,TOBOuvrage);
    end else
    begin
    	TraiteLigneTB (TOBBTB,TOBL);
    end;
  end;

end;

procedure TOF_BTRESULTCHADET.DefiniPrevuDetail (TOBPiece: TOB ; NaturePiece : string );
var TOBLignes,TOBOuvrages : TOB;
begin
	TOBLignes := TOB.Create ('LA PIECE',nil,-1);
  TOBLignes.Dupliquer (TOBPiece,false,true);
  TOBOuvrages := TOB.Create ('LES OUVRAGES',nil,-1);

  ChargelesLignesTB (TOBLIgnes,TOBPiece);
  ChargelesOuvragesTB (TOBOuvrages,TOBLignes);

  TraiteLaPieceTB (TOBBTB,TOBLignes,TOBOuvrages);

  TOBLIgnes.free;
  TOBOuvrages.free;
end;


procedure TOF_BTRESULTCHADET.ChargelesInfosprevu(NaturePiece : string; WherePiece : string='');
var Req : String;
    TOBPieces: TOB;
    Indice : integer;
    Critere : string;
    RefAffaire : string;
begin

	TOBPieces := TOB.Create ('LES PIECES',nil,-1);
  RefAffaire := 'GP_AFFAIREDEVIS';

  if Naturepiece = 'DBT' then Critere := ' AND AFF_ETATAFFAIRE IN ("ACP","TER") '
  else if Naturepiece = 'AFF' then BEGIN Critere := ' AND AFF_ETATAFFAIRE IN ("ENC","TER") '; RefAffaire:='GP_AFFAIRE'; END
  else if NaturePiece = 'ETU' then Critere := ' AND AFF_ETATAFFAIRE="ACP"';

  if WherePiece = '' then
  	WherePiece := ' GP_NATUREPIECEG="'+NaturePiece+'" '+'AND GP_AFFAIRE="'+Codechantier+'"';

	Req := 'SELECT GP_NATUREPIECEG,GP_SOUCHE,GP_NUMERO,GP_INDICEG FROM PIECE '+
  			 'LEFT JOIN AFFAIRE ON AFF_AFFAIRE='+RefAffaire+' WHERE '+
  			 WherePiece +
         Critere;

  TOBPieces.LoadDetailDBFromSQL  ('PIECE',Req,false);

  for indice := 0 to TOBPieces.detail.count -1 do
  begin
    DefiniPrevuDetail (TOBPIeces.detail[Indice],NaturePiece);
  end;

  TOBPIeces.free;
end;


procedure TOF_BTRESULTCHADET.DispatchMontants(TOBTMP, TOBECLATE: TOB);
var NatureMouv : string;
		MontantPv, Montant : double;
begin
  NatureMouv    := TOBECLATE.GetString('NATUREMOUV');
  MontantPV     := TOBECLATE.GetValue('VENTE');
  //
  if (TOBECLATE.GetString('BCO_NATUREPIECEG') = '') or     // consos saisies
     (TOBECLATE.GetString('BCO_NATUREPIECEG') = 'AFS') or   // correction BRL le 06/04/05 prise ne compte des avoirs en consos
     (TOBECLATE.GetString('BCO_NATUREPIECEG') = 'AF') or    // ...
     (TOBECLATE.GetString('BCO_NATUREPIECEG') = 'BFC') or   // Retour chantier
     (TOBECLATE.GetString('BCO_NATUREPIECEG') = 'LBT') then // consos issues de bons de livraison clients
  begin
    if NatureMouv = 'FOU' then
    begin
    	// Répartition Fournitures
    	TOBTMP.putValue('REALISEPA',TOBECLATE.getValue('ACHAT'));
    	TOBTMP.putValue('REALISEPR',TOBECLATE.getValue('REVIENT'));
    	TOBTMP.putValue('REALISEPV',TOBECLATE.getValue('VENTE'));
    	TOBTMP.putValue('QTEREALISE',TOBECLATE.getValue('BCO_QUANTITE'));
    end Else if NatureMouv = 'MO' then
    begin
    	// Répartition MO Interne
    	TOBTMP.putValue('REALISEPA',TOBECLATE.getValue('ACHAT'));
    	TOBTMP.putValue('REALISEPR',TOBECLATE.getValue('REVIENT'));
    	TOBTMP.putValue('REALISEPV',TOBECLATE.getValue('VENTE'));
    	TOBTMP.putValue('QTEREALISE',TOBECLATE.getValue('BCO_QUANTITE'));
    end Else if NatureMouv = 'FRS' then
    begin
      // Répartition Frais
    	TOBTMP.putValue('REALISEPA',TOBECLATE.getValue('ACHAT'));
    	TOBTMP.putValue('REALISEPR',TOBECLATE.getValue('REVIENT'));
    	TOBTMP.putValue('REALISEPV',TOBECLATE.getValue('VENTE'));
    	TOBTMP.putValue('QTEREALISE',TOBECLATE.getValue('BCO_QUANTITE'));
    end Else if NatureMouv = 'RES' then
    begin
      // Répartition Matériels + Outils
    	TOBTMP.putValue('REALISEPA',TOBECLATE.getValue('ACHAT'));
    	TOBTMP.putValue('REALISEPR',TOBECLATE.getValue('REVIENT'));
    	TOBTMP.putValue('REALISEPV',TOBECLATE.getValue('VENTE'));
    	TOBTMP.putValue('QTEREALISE',TOBECLATE.getValue('BCO_QUANTITE'));
    end Else if NatureMouv = 'EXT' then
    begin
      // Répartition Intérim + Sous-traitance + Location + Autres
    	TOBTMP.putValue('REALISEPA',TOBECLATE.getValue('ACHAT'));
    	TOBTMP.putValue('REALISEPR',TOBECLATE.getValue('REVIENT'));
    	TOBTMP.putValue('REALISEPV',TOBECLATE.getValue('VENTE'));
    	TOBTMP.putValue('QTEREALISE',TOBECLATE.getValue('BCO_QUANTITE'));
    end Else if NatureMouv = 'RAN' Then
    Begin
      // Répartition recettes annexes
      MontantPV := MontantPV * -1;
      TOBTMP.PutValue('RECETTEANNEXE', MontantPV);
  		TOBTMP.PutValue('NATURE','RAN : Recettes Annexes');

    end Else If NatureMouv = 'FAN' then
    Begin
      // Répartition coûts annexes
      TOBTMP.PutValue('COUTANNEXE', MontantPV);
  		TOBTMP.PutValue('NATURE','CAN : Couts Annexes');
    end;
  end else if (TOBECLATE.GetString('BCO_NATUREPIECEG') = 'BLF') or
              (TOBECLATE.GetString('BCO_NATUREPIECEG') = 'LFR')  or
              (TOBECLATE.GetString('BCO_NATUREPIECEG') = 'FF')  then
  begin
       // consos issues de bons de réception fournisseurs (engagé chantier)
     if (TOBECLATE.getValue('BCO_TRANSFORME') = '-') and (TOBECLATE.getValue('BCO_TRAITEVENTE') = '-') then
       TOBTMP.PutValue('ENGAGECHA', TOBECLATE.getValue('ACHAT'));
  end;

  if (TOBECLATE.GetString('BCO_NATUREPIECEG') = 'CF') or (TOBECLATE.GetString('BCO_NATUREPIECEG') = 'BLF')or
     (TOBECLATE.GetString('BCO_NATUREPIECEG') = 'CFR') or (TOBECLATE.GetString('BCO_NATUREPIECEG') = 'LFR')then
  begin
    if (TOBECLATE.getValue('BCO_TRANSFORME') = '-') and (TOBECLATE.getValue('GP_VIVANTE') = 'X') then
    begin
      TOBTMP.PutValue('FACTNONPAR', TOBECLATE.getValue('ACHAT'));
    end;
  end;

end;

Procedure TOF_BTRESULTCHADET.ChargementEngageFournisseur(TOBTMP,TobProv : TOB);
var Montant : double;
begin

  Montant := TobProv.GetDouble('ACHAT');

  if Montant <> 0 then
  begin
    TOBTMP.PutValue('ENGAGEFOU', TOBTMP.GetValue('ENGAGEFOU') + Montant);
    TOBTMP.PutValue('ENGAGECHA', TOBTMP.GetValue('ENGAGECHA') + Montant);
  end;

end;

procedure TOF_BTRESULTCHADET.SetChampsCommun(TOBTMP, OneTOB: TOB);

begin
  TOBTMP.PutValue('CODEARTICLE',OneTOB.getValue('CODEARTICLE')+' : '+OneTOB.getValue('LIBELLEART'));

  if (not varisnull(OneTOB.getValue('NATUREPRESTATION'))) and (OneTOB.getValue('NATUREPRESTATION')<>'') then
  begin
  	TOBTMP.PutValue('NATURE',OneTOB.getValue('NATUREPRESTATION')+' : '+OneTOB.getValue('LIBNATURE'));
  end else
  begin
  	TOBTMP.PutValue('NATURE','FOU : Fourniture');
  end;

  if OneTOB.getValue('FAMILLE1')<> '' then
  begin
  	TOBTMP.PutValue('FAMILLE1',OneTOB.getValue('FAMILLE1')+' : '+OneTOB.getValue('LIBFAMILLE1'));
  end else
  begin
  	TOBTMP.PutValue('FAMILLE1','Non définie');
  end;

  if OneTOB.getValue('FAMILLE2')<> '' then
  begin
  	TOBTMP.PutValue('FAMILLE2',OneTOB.getValue('FAMILLE2')+' : '+OneTOB.getValue('LIBFAMILLE2'));
  end else
  begin
  	TOBTMP.PutValue('FAMILLE2','Non définie');
  end;

  if OneTOB.getValue('FAMILLE2')<> '' then
  begin
  	TOBTMP.PutValue('FAMILLE3',OneTOB.getValue('FAMILLE3')+' : '+OneTOB.getValue('LIBFAMILLE3'));
  end else
  begin
  	TOBTMP.PutValue('FAMILLE3','Non définie');
  end;

end;



procedure TOF_BTRESULTCHADET.ChargelesInfosprevuAvanc;
var req : string;
    MontantPA, MontantPr,MontantPV, AvancePa,AvancePr,AvancePv,TpsPrevu,TpsAvance : double;
    TOBTMp,OneTOB : TOB;
    indice : integer;
begin
	OneTOB := TOB.Create ('ELS LIGNES',nil,-1);

  // Récupération du prévu
  Req := 'SELECT GL_CODEARTICLE AS CODEARTICLE,BNP_TYPERESSOURCE AS NATUREPRESTATION,'+
  			 'GL_LIBELLE AS LIBELLEART,'+
         'GA_FAMILLENIV1 AS FAMILLE1, GA_FAMILLENIV2 AS FAMILLE2, GA_FAMILLENIV3 AS FAMILLE3, ' +
         '(SELECT CC_LIBELLE FROM '+NomChoixCod2+' WHERE CC_TYPE="TRE" AND CC_CODE=BNP_TYPERESSOURCE) AS LIBNATURE, '+
        '(SELECT CC_LIBELLE FROM '+NomChoixCod+' WHERE CC_TYPE="FN1" AND CC_CODE=GA_FAMILLENIV1) AS LIBFAMILLE1, '+
        '(SELECT CC_LIBELLE FROM '+NomChoixCod+' WHERE CC_TYPE="FN2" AND CC_CODE=GA_FAMILLENIV2) AS LIBFAMILLE2, '+
        '(SELECT CC_LIBELLE FROM '+NomChoixCod+' WHERE CC_TYPE="FN3" AND CC_CODE=GA_FAMILLENIV3) AS LIBFAMILLE3,'+
  			 'GL_QTEFACT*GL_DPA AS ACHAT, '+
  			 'GL_QTEFACT*GL_DPR AS REVIENT, '+
         'GL_TOTALHTDEV AS VENTE, ' +
         'GL_QTEFACT AS TPS_PREVU';

  // .. et de l'avance
  Req := Req + ',GL_QTEPREVAVANC*GL_DPA AS AVANCEPA'+
  						 ',GL_QTEPREVAVANC*GL_DPR AS AVANCEPR'+
               ',GL_TOTALHTDEV*(GL_POURCENTAVANC/100.0) AS AVANCEPV'+
               ',GL_QTEPREVAVANC AS TPS_AVANCE' ;

  // .. et de la cotraitance et sous traitance
  Req := Req + ',GLC_NATURETRAVAIL' ;

  // .. suite et faim
  Req := Req + ' FROM LIGNE ' +
  						 ' LEFT JOIN ARTICLE ON GA_ARTICLE=GL_ARTICLE '+
               ' LEFT JOIN NATUREPREST ON BNP_NATUREPRES=GA_NATUREPRES '+
               ' LEFT JOIN LIGNECOMPL ON GLC_NATUREPIECEG=GL_NATUREPIECEG ' +
               '  AND GLC_SOUCHE=GL_SOUCHE AND GLC_NUMERO=GL_NUMERO ' +
               '  AND GLC_INDICEG=GL_INDICEG AND GLC_NUMORDRE=GL_NUMORDRE ';

  // .. sans oublier les conditions
  Req := req + 'WHERE GL_NATUREPIECEG = "PBT" AND GL_TYPELIGNE LIKE "AR%" AND GL_AFFAIRE="' + Codechantier + '"';

  OneTOB.LoadDetailDBFromSQL ('LIGNE',Req);
  For Indice := 0 to OneTOB.detail.count -1 do
  begin
    TOBTMP := TOB.create('UNE LIGNE',TOBBTB,-1);
    AjouteChampSup (TOBTmp);
    SetChampsCommun (TOBTMP,OneTOB.detail[Indice]);

    MontantPA := OneTOB.detail[Indice].getValue('ACHAT');
    MontantPR := OneTOB.detail[Indice].getValue('REVIENT');
    MontantPV := OneTOB.detail[Indice].getValue('VENTE');
    AvancePA  := OneTOB.detail[Indice].getValue('AVANCEPA');
    AvancePR  := OneTOB.detail[Indice].getValue('AVANCEPR');
    AvancePV  := OneTOB.detail[Indice].getValue('AVANCEPV');
    TpsPrevu  := OneTOB.detail[Indice].getValue('TPS_PREVU');
    TpsAvance := OneTOB.detail[Indice].getValue('TPS_AVANCE');

    TOBTMP.PutValue('PREVUPBTPA', MontantPA);
    TOBTMP.PutValue('PREVUPBTPR', MontantPR);
    TOBTMP.PutValue('PREVUPBTPV', MontantPV);

    TOBTMP.PutValue('AVANCEPA', AvancePa);
    TOBTMP.PutValue('AVANCEPR', AvancePr);
    TOBTMP.PutValue('AVANCEPV', AvancePV);

    TOBTMP.PutValue('QTEPREVUPBT', TpsPrevu);
    TOBTMP.PutValue('QTEAVANCE',   TpsAvance);
  end;

  OneTOB.free;

end;

Procedure TOF_BTRESULTCHADET.ChargelesInfosEngage;
Var Req       : String;
    TOBPieces : TOB;
    TOBTMp    : TOB;
    Indice    : integer;
begin

  TOBPieces := TOB.Create ('LES ENGAGEES',nil,-1);

  //FV1 : 27/06/2014 - FS#911 - MATFOR - problème de montant en résultat de chantier
  // Cumul des réceptions fournisseurs et factures issues de commandes et hors période : donc à ajouter
  Req := 'SELECT GL_QTEFACT*GL_PUHTNETDEV AS ACHAT FROM LIGNE ';
  Req := Req + ' WHERE GL_NATUREPIECEG IN ("BLF", "LFR", "FF")';
  Req := Req + ' AND GL_TYPELIGNE="ART"';
  Req := Req + ' AND GL_AFFAIRE="' + Codechantier + '" ';
  Req := Req + ' AND GL_PIECEPRECEDENTE LIKE "%;CF;%"';
  Req := Req + ' AND GL_DATEPIECE > "' + UsDateTime(DateMvtFin) + '" ';

  TOBPieces.LoadDetailDBFromSQL  ('PIECE',Req,false);
  //
  for indice := 0 to TOBPieces.detail.count -1 do
  begin
    //
    TOBTMP := TOB.create('UNE LIGNE',TOBBTB,-1);
    AjouteChampSup (TOBTmp);
    SetChampsCommun (TOBTMP,TOBPieces.detail[Indice]);
    //
    ChargementEngageFournisseur(TOBTMP,TOBPieces.detail[indice]);
  end;
  //
  TOBPIeces.free;

end;

procedure TOF_BTRESULTCHADET.ChargeLeFactured;
var req,WherePiece : string;
		TOBPieces : TOB;
    indice : integer;
begin
	TOBPieces := TOB.Create ('LES FACTURES',nil,-1);
  // Récupération du facturé
  Req := 'SELECT GP_NATUREPIECEG,GP_SOUCHE,GP_NUMERO,GP_INDICEG FROM PIECE ';
  WherePiece := '(GP_NATUREPIECEG = "FBT" OR GP_NATUREPIECEG = "ABT" OR GP_NATUREPIECEG = "FAC" OR GP_NATUREPIECEG = "AVC")';

  req := req + 'WHERE '+ WherePiece;
  Req := Req + '  AND GP_AFFAIRE="' + Codechantier + '"';
  Req := Req + '  AND GP_DATEPIECE BETWEEN "' + UsDateTime(DateMvtDeb) + '"';
  Req := Req + '  AND "' + UsDateTime(DateMvtFin) + '" ';
  TOBPieces.LoadDetailDBFromSQL  ('PIECE',Req,false);

  for indice := 0 to TOBPieces.detail.count -1 do
  begin
    DefiniPrevuDetail (TOBPIeces.detail[Indice],TOBPIeces.detail[Indice].getValue('GP_NATUREPIECEG'));
  end;

  TOBPIeces.free;
end;

procedure TOF_BTRESULTCHADET.DateTransfoOK (TOBEclate : TOB; var Montant : Double);
var Q : Tquery;
    DateTransfo : TDateTime ;
begin
  Montant := 0;
  Q := OpenSQL('SELECT BCO_DATEMOUV, BCO_MONTANTACH FROM CONSOMMATIONS WHERE BCO_TRANSFORME="-" AND BCO_LIENTRANSFORME='+FloatToStr(TOBEclate.GetValue('BCO_NUMMOUV')), True);
  while Not Q.EOF do
  begin
  	DateTransfo := StrToDate(Q.FindField('BCO_DATEMOUV').AsString);
    if not ((DateTransfo >= DateMvtDeb) and (DateTransfo <= DateMvtFin)) then
      Montant := Montant + Q.FindField('BCO_MONTANTACH').AsFloat;
    Q.Next;
  end;
	Ferme (Q);
end;


procedure TOF_BTRESULTCHADET.RechTiersFou(CodeFrs : String; Var LibTiers :String);
var ferr     : integer;
    StSQL    : string;
begin

	if not GetLibTiers('FOU',CodeFrs,LibTiers) then LibTiers := 'Fournisseur inexistant';

end;


Initialization
  registerclasses ( [ TOF_BTRESULTCHADET ] ) ;
end.
