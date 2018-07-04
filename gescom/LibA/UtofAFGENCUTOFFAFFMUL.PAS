{***********UNITE*************************************************
Auteur  ...... : PL
Créé le ...... : 20/08/2002
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : AFGENCUTOFFAFFMUL ()
Mots clefs ... : TOF;AFGENCUTOFFAFFMUL
*****************************************************************}
Unit UtofAFGENCUTOFFAFFMUL;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFDEF EAGLCLIENT}
   eMul,Maineagl,
{$ELSE}
   Mul,AssistCreationAffaire,dbTables,HDB,FE_Main,
{$ENDIF}
   ComCtrls,paramsoc,
   forms,sysutils,HCtrls,Hstatus,HEnt1, M3FP, HTB97,
   HMsgBox, UTOF,Utob,DicoAf,SaisUtil,AffaireUtil,utilcutoff,
   Ent1,UtilPgi,uTofAfBaseCodeAffaire, TraducAffaire,uTofAfSaisDate,
   UtilMulTrt,UtofAFFormule;

Type
  TOF_AFGENCUTOFFAFFMUL = Class (TOF_AFBASECODEAFFAIRE)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnClose                  ; override ;
    procedure NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2, Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers, Tiers_:THEdit); override ;
    procedure GenereLesCutoffParAffaire ;
    function GenerLesCutOff(TobAffairesSelec:TOB):integer ;
  end ;

Procedure AFLanceFiche_AfGenCutOffAFF_Mul(Range:string);

Implementation

procedure TOF_AFGENCUTOFFAFFMUL.OnArgument (S : String ) ;
begin
  Inherited ;
end ;

procedure TOF_AFGENCUTOFFAFFMUL.OnUpdate ;
begin
  Inherited ;
  {$IFDEF EAGLCLIENT}
  TraduitAFLibGridSt(TFMul(Ecran).FListe);
  {$ELSE}
  TraduitAFLibGridDB(TFMul(Ecran).FListe);
  {$ENDIF}

  If not(ctxScot in V_PGI.PGIContexte)  then
  begin
      SetControlVisible ('TT_MOISCLOTURE' , False);
      SetControlVisible ('T_MOISCLOTURE' , False);
      SetControlVisible ('TT_MOISCLOTURE_' , False);
      SetControlVisible ('T_MOISCLOTURE_' , False);
  end;

end ;

procedure TOF_AFGENCUTOFFAFFMUL.NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2, Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers, Tiers_:THEdit);
begin
Aff:=THEdit(GetControl('AFF_AFFAIRE'));
Aff0:=THEdit(GetControl('AFF_AFFAIRE0'));
Aff1:=THEdit(GetControl('AFF_AFFAIRE1'));
Aff2:=THEdit(GetControl('AFF_AFFAIRE2'));
Aff3:=THEdit(GetControl('AFF_AFFAIRE3'));
Aff4:=THEdit(GetControl('AFF_AVENANT'));
Aff_:=THEdit(GetControl('AFF_AFFAIRE_'));
Aff0_:=THEdit(GetControl('AFF_AFFAIRE0_'));
Aff1_:=THEdit(GetControl('AFF_AFFAIRE1_'));
Aff2_:=THEdit(GetControl('AFF_AFFAIRE2_'));
Aff3_:=THEdit(GetControl('AFF_AFFAIRE3_'));
Aff4_:=THEdit(GetControl('AFF_AVENANT_'));
Tiers:=THEdit(GetControl('AFF_TIERS'));
Tiers_:=THEdit(GetControl('AFF_TIERS_'));
end;

procedure TOF_AFGENCUTOFFAFFMUL.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_AFGENCUTOFFAFFMUL.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_AFGENCUTOFFAFFMUL.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_AFGENCUTOFFAFFMUL.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_AFGENCUTOFFAFFMUL.GenereLesCutoffParAffaire ;
var
TobMAff : TOB;
iNbAffaire:integer;
begin
if (trim(GetParamSoc('SO_AFFORMULCUTOFF'))='') then
begin
  PGIInfoAf('Pour faire une génération, vous devez saisir la formule de calcul du Cut Off.', Ecran.Caption);
  AFLanceFiche_FormuleCutOff ('');
  if (trim(GetParamSoc('SO_AFFORMULCUTOFF'))='') then
  begin
    if TFMul(Ecran).FListe.AllSelected then
      TFMul(Ecran).FListe.AllSelected:=False
    else TFMul(Ecran).FListe.ClearSelected;
    TFMul(Ecran).bSelectAll.Down := False ;
    exit;
  end;
end;
TobMAff := TOB.Create ('', Nil, -1);
try
SourisSablier;
// Pl le 22/11/02 : on travail sur la vue complète pour récupérer les champs qui seront utilisés dans
// la formule de calcul du cutoff et qui ne sont pas forcément définis (SELECT *)
//TraiteEnregMulListe (TFMul(Ecran), 'AFF_AFFAIRE','AFAFFTIERS', TobMAff, True);
TraiteEnregMulTable (TFMul(Ecran), 'SELECT * FROM AFAFFTIERS', 'AFF_AFFAIRE', 'AFAFFTIERS', 'AFF_AFFAIRE',
                      'AFAFFTIERS', TobMAff, True);

iNbAffaire := TobMAff.Detail.count;

if PGIAskAF ('Attention, suppression de tous les cutoff déjà générés à cette date, pour les affaires sélectionnées.'+chr(10)+'Confirmez-vous la génération des cut off ?', Ecran.Caption) <> mrYes then
    exit;

// Generation des cut off pour toutes les affaires selectionnées
if (iNbAffaire<>0) then
    GenerLesCutOff (TobMAff)
else
    PGIInfoAf('Vous devez sélectionner au-moins une affaire', Ecran.Caption);


// Rafraichit la liste
TFMul(Ecran).ChercheClick;

finally
TobMAff.Free;
SourisNormale;
end;

TToolBarButton97(GetControl('bSelectAll')).Down:=false;
end;

function TOF_AFGENCUTOFFAFFMUL.GenerLesCutOff (TobAffairesSelec : TOB) : integer;
  var
  //iNbAffaires:integer;
  iPosDateDeb : integer;
  RepCO, Arg : string;
  sDateCO, sDateDeb, sDateFin : string;
  repGener, bAvecCutOff : boolean;
begin
  repGener := false;
  bAvecCutOff := true;
  Result := 0;
  sDateCO := ''; RepCO := '';


  try
    Arg := 'TEST:X;TITRE:Saisie de la date de Cut off;LIBELLE:Date du Cut off;ZORI:CUTOFFAFF;ZZDATE:'
//            + datetostr (GetParamSoc ('SO_DATECUTOFF')) + ';ZZDATEDEB:' + datetostr (GetParamSoc ('SO_DATECUTOFF'));
//            + datetostr (GetParamSoc ('SO_DATECUTOFF')) + ';ZZDATEDEB:01/01/1900';
            + datetostr (GetParamSoc ('SO_DATECUTOFF')) + ';ZZDATEDEB:' + datetostr (GetParamSoc ('SO_AFDATEDEBCAB'));

    // Recuperation de la date de cut off globale pour toutes les affaires selectionnees
    sDateCO := AFLanceFiche_SaisieDate (Arg);

    if (sDateCO <> '0') then
      begin
        sDateFin := sDateCO;
        sDateDeb := '01/01/1900';
        if (pos ('SANSCUTOFF', sDateCO) <> 0) then
          begin
            bAvecCutOff := false;
            sDateCO := copy (sDateCO, 1, pos ('SANSCUTOFF', sDateCO) - 2);
          end;
        iPosDateDeb := pos (';', sDateCO);
        if iPosDateDeb <> 0 then
          begin
            sDateFin := copy (sDateCO, 1, iPosDateDeb - 1);
            sDateDeb := copy (sDateCO, iPosDateDeb + 1, length (sDateCO));
            if (pos ('SANSCUTOFF', sDateCO) <> 0) then bAvecCutOff := false;
          end;

        repGener := BoucleCutOff (false, strtodate (sDateDeb), strtodate (sDateFin), TobAffairesSelec, bAvecCutOff);
      end;

  except
    Result := 1;
  end;

  if not repGener then Result := 2;
end;

procedure AGLGenereLesCutOff( parms: array of variant; nb: integer );
var  F : TForm;
     LaTof : TOF;
begin
F:=TForm(Longint(Parms[0]));
if (F is TFMul) then Latof:=TFMul(F).Latof else laTof := nil;
if (Latof is TOF_AFGENCUTOFFAFFMUL) then TOF_AFGENCUTOFFAFFMUL(LaTof).GenereLesCutoffParAffaire else exit;
end;

Procedure AFLanceFiche_AfGenCutOffAFF_Mul(Range:string);
begin
AGLLanceFiche ('AFF','AFGENCUTOFFAFFMUL',Range,'','');
end;

Initialization
  registerclasses ( [ TOF_AFGENCUTOFFAFFMUL ] ) ;
  RegisterAglProc( 'GenereLesCutOff',True,0,AGLGenereLesCutOff);

end.
