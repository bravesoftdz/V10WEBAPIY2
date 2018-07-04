{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 03/07/2002
Modifié le ... : 29/11/2002
Description .. : Source TOF de la FICHE : AFMODIFCUTOFF_MUL ()
Mots clefs ... : TOF;AFMODIFCUTOFF_MUL
*****************************************************************}
Unit UtofAFMODIFCUTOFF_MUL ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFDEF EAGLCLIENT}
     Maineagl,emul,UtileAGL,
{$ELSE}
     db,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}FE_Main,mul,EdtREtat,
{$ENDIF}
{$IFDEF BTP}
	 CalcOleGenericBTP,
{$ENDIF}
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox, Grids,ExtCtrls,Windows,
     UTOF, M3FP,Hqry,ed_tools, 
     UtilMulTrt, utob, dicobtp, htb97,dialogs,utilcutoff,
     UTofAfBaseCodeAffaire,UtofAFMODIFCUTOFF,Affaireutil,SaisUtil,ParamSoc,AppreciationUtil,ActiviteUtil ;

Type
  TOF_AFMODIFCUTOFF_MUL = Class (TOF_AFBASECODEAFFAIRE)
  private
    GS : THGRID;
    TobAfcumul :TOB;
    TypeCumul,TypeSaisie,LesColonnes,Titre :String;
    DEV : RDEVISE;
    FindLigne: TFindDialog;
    FindDebut,ModifGlobal,Erreur :boolean;
    LigOld : integer;
    PCUMULSAI : TPanel;
    TOTFAE,TOTPCA,TOTAAE : THNumEdit;
    //Saisie des montants FAE AAE PCA en Global par affaire
    procedure GSInitSaisieCutOff;
    procedure GSChargeSaisieCutOff;
    procedure GSCellEnter (Sender: TObject; var ACol,ARow: integer; var Cancel: boolean);
    procedure GSCellExit (Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure GSRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    procedure GSEnter(Sender: TObject);
    procedure GSExit(Sender: TObject);
    procedure GSDblClick(Sender: TObject);
    procedure GSEnregistreClick(Sender: TObject);
    procedure GSEnregistre;
    procedure BEtatImprimerClick(Sender: TObject);
//    procedure GSCalculMontant(Montant:Double;Repartition:integer;ChampMt:string;TOBL :TOB);
    procedure GSChercherClick(Sender: TObject);
    procedure GSFindLigne(Sender: TObject);
    procedure GSColumnWidthsChanged(Sender: TObject);
    procedure GSChargeTotaux (var CTOTAL :THNumEdit; CHTOTAL,CHCUMUL :string; ColTOTAL :integer);
  public
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (stArgument : String ) ; override ;
    procedure OnClose                  ; override ;
    procedure NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2, Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers, Tiers_:THEdit);override ;
    procedure AfModifCutOff;
    function ModifLesCutOff(TobAffairesSelec:TOB):integer;
  END ;

Const ColObjects   : integer = 0;
      ColAff1      : integer = 1;
      ColAff2      : integer = 2;
      ColAff3      : integer = 3;
      ColAven      : integer = 4;
      ColTiers     : integer = 5;
      ColLibelle   : integer = 6;
      ColFAE       : integer = 7;
      ColPCA       : integer = 8;
      ColAAE       : integer = 9;

TexteMsgAffaire: array[1..2] of string	= (
  {1}        'Confirmez-vous l''enregistrement des montants modifiés ? ',
  {2}        'Enregistrer les montants modifiés');

Procedure AFLanceFiche_ModifCutOffMul(Argument:String);

Implementation

procedure TOF_AFMODIFCUTOFF_MUL.OnArgument (stArgument : String ) ;
var Critere, Champ, valeur, PArtie0, Partie1,Partie2, Partie3, Avenant: String;
    X : integer;
    Date : TdateTime;
begin
  Erreur :=False;
  LigOld := -1;
  Critere := (Trim(ReadTokenSt(stArgument)));
  Date := Idate1900;
  While (Critere <>'') do
  begin
    X := pos (':', Critere);
    if (X <> 0) then
       begin
       Champ:=copy(Critere,1,X-1);
       Valeur:=Copy (Critere,X+1,length(Critere)-X);
       end;
    if (Champ = 'TYPE')   then TypeCumul:=valeur;
    if (Champ = 'SAISIE') then TypeSaisie:=valeur;
    if (Champ = 'AFFAIRE') then
      begin // mcd 08/04/03 pour appel dpeuis TB
	{$IFDEF BTP}
      BTPCodeAffaireDecoupe(valeur,Partie0,Partie1,Partie2,Partie3,Avenant,TaCOnsult,False);
    {$ELSE}
      CodeAffaireDecoupe(valeur,Partie0,Partie1,Partie2,Partie3,Avenant,TaCOnsult,False);
    {$ENDIF}
      SetControLText ('ACU_AFFAIRE',Valeur);
        If Valeur <>'' then
        begin
        SetControLText ('ACU_AFFAIRE1',Partie1);
        SetControLText ('ACU_AFFAIRE2',Partie2);
        SetControLText ('ACU_AFFAIRE3',Partie3);
        SetControLText ('ACU_AVENANT',Avenant);
        SetControlEnabled ('ACU_TIERS',False);
        SetControlEnabled ('ACU_AFFAIRE1',False);
        SetControlEnabled ('ACU_AFFAIRE2',False);
        SetControlEnabled ('ACU_AFFAIRE3',False);
        SetControlEnabled ('ACU_AVENANT',False);
        setControlEnabled ('BSELECTAFF1',False);
        setControlEnabled ('BEFFACEAFF1',False);
        end;
      BchangeTiers :=False;
      end;
    if (Champ = 'DATE') then Date:=StrTodate(valeur);
    if (Champ = 'TIERS') then
      begin
      SetControlText('ACU_TIERS',valeur);
      SetControlEnabled ('ACU_TIERS',False);
      end;
    Critere:=(Trim(ReadTokenSt(stArgument)));
  end;
Inherited ;
  if GetParamSoc('SO_DATECUTOFF')=0 then   //AB-20031022
    SetControlText('ACU_DATE',DateToSTr(iDate1900))
  else
    SetControlText('ACU_DATE',DateToSTr(GetParamSoc('SO_DATECUTOFF')));

  if Date <>Idate1900 then SetControlText('ACU_DATE',DateToSTr(Date));   //mcd 08/04/2003
  if (TypeSaisie='GLOBAL') then GSInitSaisieCutOff;
end ;

procedure TOF_AFMODIFCUTOFF_MUL.OnLoad ;
begin
  Inherited ;
  SetControlText('ACU_TYPEAC',TypeCumul);
  if (TypeSaisie='GLOBAL') then GSChargeSaisieCutOff;
 (*   // mcd 26/05/03 ajout test % date arrete période... mais tout est bloqué
      // faut il faire le test, mais permettre si date cut off = date dernier cut off
      // et on ne peut même plus consulter les anciens cut off....
 if (GetParamSoc('SO_AFDateDebutAct')<>0) and (GetParamSoc('SO_AFDateDebutAct')<>iDate1900) and (GetParamSoc('So_AFDateDebutAct')<>iDate2099) then
    begin
    if (StrToDAte(GetControLtext ('ACU_DATE')) <GetParamSoc('SO_AFDateDebutAct')) then
         begin
         PGIInfo ('la date saisie est inférieure à la date de début d''activité', titre);
         LastError :=1;
         Erreur :=True;
         exit;
         end;
    end;  *)

end ;

procedure TOF_AFMODIFCUTOFF_MUL.NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2, Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers, Tiers_:THEdit);
begin
Aff:=THEdit(GetControl('ACU_AFFAIRE'));
Aff0:=nil;
Aff1:=THEdit(GetControl('ACU_AFFAIRE1'));
Aff2:=THEdit(GetControl('ACU_AFFAIRE2'));
Aff3:=THEdit(GetControl('ACU_AFFAIRE3'));
Aff4:=THEdit(GetControl('ACU_AVENANT'));
Tiers:=THEdit(GetControl('ACU_TIERS'));
end;

procedure TOF_AFMODIFCUTOFF_MUL.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_AFMODIFCUTOFF_MUL.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_AFMODIFCUTOFF_MUL.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_AFMODIFCUTOFF_MUL.OnClose ;
begin
  Inherited ;
  If Erreur then
    begin
      LastError := (-1);
      LastErrorMsg := '';
      Erreur := false;
      exit;     // il y eu une erreur, on reste sur la même fiche
    end;
  if (TypeSaisie='GLOBAL') then
  begin
   if ModifGlobal and (PGIAskAF (TexteMsgAffaire[1], Titre)=mrYes) then GSEnregistre;
   TobAfcumul.cleardetail; TobAfcumul.free; TobAfcumul:= NIL;
   FindLigne.Destroy;
  end;
end ;

procedure TOF_AFMODIFCUTOFF_MUL.AfModifCutOff;
var
TobMAff :TOB;
iNbAffaire:integer;
begin

TobMAff:=TOB.Create('',Nil,-1);
try
SourisSablier;
TraiteEnregMulListe (TFMul(Ecran), 'ACU_AFFAIRE','AFCUTOFF', TobMAff, True);

iNbAffaire:=TobMAff.Detail.count;

//
if (iNbAffaire<>0) then
    ModifLesCutOff(TobMAff)
else
    PGIInfoAf('Vous devez sélectionner au-moins une affaire', Ecran.Caption);

(*if (RepCutOff=0) then
   PgiInfoAf('Les modifications  été effectuée avec succès', Ecran.Caption)
else
   PgiInfoAf('Un problème est survenu lors de la revalorisation', Ecran.Caption);*)
// Rafraichit la liste
TFMul(Ecran).ChercheClick;

finally
TobMAff.Free;
SourisNormale;
end;

TToolBarButton97(GetControl('bSelectAll')).Down:=false;
end;


function TOF_AFMODIFCUTOFF_MUL.ModifLesCutOff(TobAffairesSelec:TOB):integer;
var
i,iNbAffaires:integer;
RepCO, Arg:string;
TOBTiers, LesTOBTiers : TOB;
RepBlocage : T_TypeBlocAff;
dDateDebItv, dDateFinItv : TDateTime;
begin
Result:=0;

if DetermineModeEclatCutOff=tmeSans then exit;

TOBTiers := TOB.Create('TIERS',Nil,-1) ;
LesTOBTiers := TOB.Create('Les TOB Tiers',Nil,-1) ;
try
//
iNbAffaires := TobAffairesSelec.Detail.Count;
try
i:=0;
while (i < iNbAffaires) and (RepCO<>'STOP') do
    begin
    Arg := 'AFFAIRE=' + TobAffairesSelec.Detail[i].GetValue('ACU_AFFAIRE') + ';';
    Arg := Arg + 'TIERS=' + TobAffairesSelec.Detail[i].GetValue('ACU_TIERS')+ ';';
    Arg := Arg + 'DATE=' + DateToStr(TobAffairesSelec.Detail[i].GetValue('ACU_DATE')) + ';';
    if RemplirTOBTiersP ( TOBTiers, LesTOBTiers, TobAffairesSelec.Detail[i].GetValue('ACU_TIERS') ) then
      Arg := Arg + 'MOISCLOTURE=' + inttostr(TOBTiers.GetValue('T_MOISCLOTURE')) + ';';
    RepBlocage := BlocageAffaire( 'CUM', TobAffairesSelec.Detail[i].GetValue('ACU_AFFAIRE'), V_PGI.groupe,
                                  V_PGI.DateEntree, false, true, false, dDateDebItv, dDateFinItv, nil);
    if (RepBlocage <> tbaAucun) then Arg := Arg + 'BLOCAGE=OUI;';

    RepCO:=AFLanceFiche_ModifCutOff(Arg) ;
    Inc(i);
    end;

finally
TOBTiers.Free;
LesTOBTiers.Free;
end;

except
Result:=1;
end;
end;

/////////////// Saisie des montants FAE AAE PCA en Global par affaire //////////////
/////////////// Un Grid de saisie GS remplace Fliste du Mul           //////////////

Procedure TOF_AFMODIFCUTOFF_MUL.GSInitSaisieCutOff;
begin
  Ecran.Caption := traduitGA('Modification Globale');
  updatecaption(Ecran);
  SetControlVisible('GS',True);
  SetControlVisible('Fliste',False);
  SetControlVisible('BSelectAll',False);
  SetControlVisible('BParamListe',False);
  SetControlVisible('BExport',False);
  TFMul(Ecran).Q.Liste := 'AFMULCUTOFFGLO';
  GS:=THGRID(GetControl('GS'));
  GS.OnCellEnter:=GSCellEnter ;
  GS.OnRowEnter:=GSRowEnter;
  GS.OnCellExit:=GSCellExit;
  GS.OnExit:=GSExit;
  GS.OnEnter:=GSEnter;
  GS.OnDblClick:=GSDblClick;
  TToolBarButton97(GetControl('BVOIRAFFAIRE')).Onclick:=GSDblClick;
  TToolBarButton97(GetControl('BOUVRIR')).Onclick:=GSEnregistreClick;
  TToolBarButton97(GetControl('BOUVRIR')).Hint:= TexteMsgAffaire[2];
  TToolBarButton97(GetControl('BRechercher')).Onclick:=GSChercherClick;
  FindLigne:=TFindDialog.Create(Ecran);
  FindLigne.OnFind:=GSFindLigne ;
  GS.ColCount:= 10;
  // PL le 22/10/03 : on met tout le monde pareil pour une question de cohérence avec la liste
  // de génération des cut off : accord GM + Anne F
//  if ctxScot In V_PGI.PGIContexte then
      LesColonnes := 'ACU_TIERS;ACU_AFFAIRE1;ACU_AFFAIRE2;ACU_AFFAIRE3;ACU_AVENANT;AFF_LIBELLE;CUMFAE;CUMPCA;CUMAAE';
//  else
//      LesColonnes := 'ACU_AFFAIRE1;ACU_AFFAIRE2;ACU_AFFAIRE3;ACU_AVENANT;ACU_TIERS;AFF_LIBELLE;CUMFAE;CUMPCA;CUMAAE';
  ColCodeAffaireGrid (GS,LesColonnes);
  GS.Cells[ColLibelle,0] := TraduitGA('Descriptif affaire');      GS.ColWidths[ColLibelle] := 100;
  GS.Cells[ColFAE,0] := TraduitGA('FAE');             GS.ColWidths[ColFAE] := 80;
  GS.ColAligns[ColFAE]:=taRightJustify;             GS.ColFormats[ColFAE]:='##0.##';
  GS.Cells[ColAAE,0] := TraduitGA('AAE');            GS.ColWidths[ColAAE] := 80;
  GS.ColAligns[ColAAE]:=taRightJustify;             GS.ColFormats[ColAAE]:='##0.##';
  GS.Cells[ColPCA,0] := TraduitGA('PCA');                GS.ColWidths[ColPCA] := 80;
  GS.ColAligns[ColPCA]:=taRightJustify;             GS.ColFormats[ColPCA]:='##0.##';
  AffecteGrid(GS,taModif) ;
  Titre := TFMul(Ecran).caption;
end;

Procedure TOF_AFMODIFCUTOFF_MUL.GSChargeSaisieCutOff;
var F : TFMul ;
    Select,StWhere, StOrderBy :string;
    QQ : TQuery;
begin
  F:=TFMul(Ecran);
  StWhere := RecupWhereCritere(F.Pages) ;
  if TobAfcumul<>Nil then begin TobAfcumul.Free; TobAfcumul:=Nil; end ;
  TobAfcumul := TOB.Create ('Cumul CutOff', Nil, -1) ;
  Select := 'SELECT ACU_TIERS,ACU_AFFAIRE,ACU_AFFAIRE1,ACU_AFFAIRE2,ACU_AFFAIRE3,ACU_AVENANT,'+
            'AFF_LIBELLE,AFF_DEVISE,ACU_TYPEAC,ACU_DATE,CUMCOO,CUMFAE,CUMAAE,CUMPCA '+
            'FROM AFCUTOFF ';
  StOrderBy := 'ORDER BY ACU_TIERS,ACU_AFFAIRE';

  QQ := nil;
  Try
    QQ := OpenSQL(Select + StWhere + StOrderBy,True) ;
    TobAfcumul.LoadDetailDB('Lignes CutOff','','',QQ,True);
    GS.VidePile(false) ;
    TobAfcumul.PutGridDetail(GS,false,false,LesColonnes,True);
    TobAfcumul.SetAllModifie (false);
    F.Hmtrad.ResizeGridColumns(GS) ;
    // PL le 29/11/02 : Ne pas forcer sur le grid : en cas de protection sur la première affaire bloquee,
    // se positionne dessus alors qu'on ne devrait pas
    //GS.SetFocus;
    //GS.Col := ColFAE; GS.Row := 1;
    TPageControl(GetControl('PAGES')).ActivePage:=TTabSheet(getcontrol('PCRITERE'));
    //SetFocusControl('ACU_AFFAIRE1');
    // Affiche le nombre de lignes et les totaux FAE PCA AAE
    PCUMULSAI :=TPANEL(GetControl('PCUMULSAI'));
    PCUMULSAI.visible := true;
    PCUMULSAI.caption := traduitGA('Totaux ('+IntToStr(TobAfcumul.detail.count)+ ' lignes)');
    GSChargeTotaux (TOTFAE,'TOTFAE','CUMFAE',ColFAE);
    GSChargeTotaux (TOTPCA,'TOTPCA','CUMPCA',ColPCA);
    GSChargeTotaux (TOTAAE,'TOTAAE','CUMAAE',ColAAE);
    GS.OnColumnWidthsChanged := GSColumnWidthsChanged;
    TToolbarButton97(getControl('BIMPRIMER')).OnClick := BEtatImprimerClick ;
  Finally
    Ferme(QQ);
  end;
end;
(*
procedure TOF_AFMODIFCUTOFF_MUL.GSCalculMontant(Montant:Double;Repartition:integer;ChampMt:string;TOBL :TOB);
begin
  if (Montant=0.0) then
  begin
    TOBL.putvalue(ChampMt,0.0); TOBL.putvalue(ChampMt+'DEV',0.0); TOBL.putvalue(ChampMt+'CON',0.0);
  end else
  begin
    Montant := Arrondi(Montant/Repartition,V_PGI.OkDecV);
    TOBL.putvalue(ChampMt,Montant);
    ConvertChampTob(ChampMt,DEV,TOBL);
  end;
end;
*)
procedure TOF_AFMODIFCUTOFF_MUL.GSEnregistreClick(Sender: TObject);
begin
  if (PGIAskAF (TexteMsgAffaire[1], Titre)<>mrYes) then  exit;
  GSEnregistre;
end;

procedure TOF_AFMODIFCUTOFF_MUL.GSEnregistre;
var i_ind1 : integer;
    TobDBAfcumul, TOBL : TOB;
    Q : TQuery;
    Select : string;
    dSommeSolde, dSommeFAE, dSommeAAE, dSommePCA : double;
begin

 InitMoveProgressForm (Ecran,'','Enregistrement et éclatement des Cutoff.',TobAfcumul.detail.count,TRUE,TRUE) ;

//  TT:=DebutProgressForm (Ecran,'','Enregistrement et éclatement des Cutoff.',TobAfcumul.detail.count,TRUE,TRUE) ;
  TobDBAfcumul := TOB.Create ('Table CutOff', Nil, -1) ;
  try
    for i_ind1 := 0 to TobAfcumul.detail.count-1 do
    begin
      TOBL := TobAfcumul.detail[i_ind1];
      if ( TOBL=nil) then continue;
       MoveCurProgressForm ('');
  // SELECT * : nombre de lignes restreint
      Select  := 'SELECT * FROM AFCUMUL WHERE ';
      Select := Select + 'ACU_TYPEAC = "'+TOBL.getvalue('ACU_TYPEAC')+'" AND ';
      Select := Select + 'ACU_DATE = "'+UsDateTime(TOBL.getvalue('ACU_DATE'))+'" AND ';
      Select := Select + 'ACU_TIERS = "'+TOBL.getvalue('ACU_TIERS')+'" AND ';
      Select := Select + 'ACU_AFFAIRE = "'+TOBL.getvalue('ACU_AFFAIRE')+'"';
      TobDBAfcumul.InitValeurs;
      Q:= OpenSQL(Select,True) ;
      if Not Q.EOF then TobDBAfcumul.LoadDetailDB('AFCUMUL','','',Q,True)
      else begin Ferme(Q); continue; end;
      Ferme(Q);
      Dev.Code := TOBL.GetValue ('AFF_DEVISE');
      if Dev.code = '' then DEV.Code:= V_PGI.DevisePivot;
      GetInfosDevise(DEV);
      // répartition en pourcentage du montant global
      dSommeSolde:=TOBL.getvalue('CUMCOO');
      dSommeFAE:=TOBL.getvalue('CUMFAE');
      dSommeAAE:=TOBL.getvalue('CUMAAE');
      dSommePCA:=TOBL.getvalue('CUMPCA');
      if (dSommeSolde<>0) then
          begin
          EclatementCutoff(dSommeSolde,dSommeFAE, dSommeAAE, dSommePCA, TobDBAfcumul, TOBL.getvalue('ACU_AFFAIRE'), '');
          TobDBAfcumul.InsertOrUpdateDB;
          end;
      (*    nbDetail := TobDBAfcumul.detail.count;
          for i_ind2:=0 to TobDBAfcumul.detail.count -1 do
          begin    // Ajouter  la répartition en pourcentage du montant global
            GSCalculMontant(TOBL.getvalue('CUMFAE'),nbDetail,'ACU_FAE',TobDBAfcumul.detail[i_ind2]);
            GSCalculMontant(TOBL.getvalue('CUMAAE'),nbDetail,'ACU_AAE',TobDBAfcumul.detail[i_ind2]);
            GSCalculMontant(TOBL.getvalue('CUMPCA'),nbDetail,'ACU_PCA',TobDBAfcumul.detail[i_ind2]);
            TobDBAfcumul.detail[i_ind2].InsertOrUpdateDB;
          end;*)
        TobDBAfcumul.cleardetail;
    end;
  finally
     FiniMoveProgressForm;
    TobDBAfcumul.free;
  end;
end;

procedure TOF_AFMODIFCUTOFF_MUL.GSExit(Sender: TObject);
var
Cancel:boolean;
ACol, ARow : integer;
begin
if csDestroying in Ecran.ComponentState then Exit ;

Cancel:=false;
ACol := GS.Col;
ARow := GS.Row;
GSCellExit(Sender, ACol, ARow, Cancel);

if Not Cancel then LigOld := -1;
end;

procedure TOF_AFMODIFCUTOFF_MUL.GSEnter(Sender: TObject);
var
ACol,ARow:integer;
Chg, Cancel : boolean;
begin
if csDestroying in Ecran.ComponentState then Exit ;

Chg:=false; Cancel:=false;
ACol := GS.Col;
ARow := GS.Row;

GSRowEnter(GS, ARow, Cancel, Chg);
if (Not Cancel) then
   GSCellEnter(GS, ACol, ARow, Cancel);

end;

procedure TOF_AFMODIFCUTOFF_MUL.GSRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
var
RepBlocage : T_TypeBlocAff;
dDateDebItv, dDateFinItv : TDateTime;
iLigne, Inc : integer;
bBloqueParle : boolean;
begin
if csDestroying in Ecran.ComponentState then Exit ;
if (TobAfcumul = nil) or (TobAfcumul.Detail.count = 0) then exit;

if Not(GoRowSelect in GS.Options) then
    // si pas de Blocage complet sur le grid, on regarde s'il y a un blocage sur l'affaire
    begin
      iLigne := GS.Row;
      bBloqueParle := true;
      if (iLigne < LigOld) and (iLigne > 1) then Inc := -1 else Inc := 1;
        repeat
          RepBlocage := BlocageAffaire( 'CUM', TobAfcumul.Detail[iLigne-1].GetValue('ACU_AFFAIRE'), V_PGI.groupe,
                                        V_PGI.DateEntree, bBloqueParle, bBloqueParle, false, dDateDebItv, dDateFinItv, nil);
          if (RepBlocage <> tbaAucun) then
            begin
            iLigne := iLigne + Inc;
            if (iLigne = 0) then begin Inc := 1; iLigne := 2; end;
            // Si on a deja eu un message bloquant, on n'en veut plus pour les autres
            bBloqueParle := false;
            end;

        until (RepBlocage = tbaAucun) or ((iLigne = 0) and (Inc = -1)) or ((iLigne >= GS.RowCount) and (Inc = 1));

        if (RepBlocage = tbaAucun) then
          GS.Row := iLigne
        else
          if (LigOld <> -1) then
            GS.Row := LigOld
          else
            begin
            //PGIInfoAf('Blocages sur affaires : saisie impossible.', Ecran.Caption);
            TPageControl(GetControl('PAGES')).ActivePage:=TTabSheet(getcontrol('PCRITERE'));
            SetFocusControl('ACU_AFFAIRE1');
            end;
    end ;

end;

procedure TOF_AFMODIFCUTOFF_MUL.GSCellEnter(Sender:TObject;var ACol,ARow:integer;var Cancel:boolean);
begin
  if csDestroying in Ecran.ComponentState then Exit ;

  if (GS.Col < ColFAE)  then  GS.Col := ColFAE ;
end;

procedure TOF_AFMODIFCUTOFF_MUL.GSCellExit(Sender:TObject;var ACol,ARow:Integer;var Cancel:Boolean);
var TOBL :TOB;
    Montant :double;
begin
  if csDestroying in Ecran.ComponentState then Exit ;
  TOBL := TOB(GS.objects[ColObjects,ARow]);
  if (TOBL=nil) then exit;
  if ((ACol=ColFAE) and (GS.Cells[ColFAE,Arow]<>'')) then
  begin
    Montant := Arrondi(Valeur(GS.Cells[ColFAE,Arow]),V_PGI.OkDecV);
    GS.Cells[ColFAE,Arow] := FloattoStr(Montant);
    if (Montant <> TOBL.getvalue('CUMFAE')) then
    begin
      ModifGlobal := true;
      TOBL.putvalue('CUMFAE',Montant);
      GSChargeTotaux (TOTFAE,'TOTFAE','CUMFAE',ColFAE);
    end;
  end else if ((ACol=ColAAE) and (GS.Cells[ColAAE,Arow]<>'')) then
  begin
    Montant := Arrondi(Valeur(GS.Cells[ColAAE,Arow]),V_PGI.OkDecV);
    GS.Cells[ColAAE,Arow] := FloattoStr(Montant);
    if (Montant <> TOBL.getvalue('CUMAAE')) then
    begin
      ModifGlobal := true;
      TOBL.putvalue('CUMAAE',Montant);
      GSChargeTotaux (TOTAAE,'TOTAAE','CUMAAE',ColAAE);
    end;
  end else if ((ACol=ColPCA) and (GS.Cells[ColPCA,Arow]<>'')) then
  begin
    Montant := Arrondi(Valeur(GS.Cells[ColPCA,Arow]),V_PGI.OkDecV);
    GS.Cells[ColPCA,Arow] := FloattoStr(Montant);
    if (Montant<>TOBL.getvalue('CUMPCA')) then
    begin
      ModifGlobal := true;
      TOBL.putvalue('CUMPCA',Montant);
      GSChargeTotaux (TOTPCA,'TOTPCA','CUMPCA',ColPCA);
    end;
  end;

if Not Cancel then
    begin
    LigOld := ARow;
    end;
end;

procedure TOF_AFMODIFCUTOFF_MUL.GSDblClick(Sender: TObject);
var CodeAffaire : string;
begin
if csDestroying in Ecran.ComponentState then Exit ;
if (TobAfcumul = nil) or (TobAfcumul.Detail.count = 0) then exit;
CodeAffaire := TOB(GS.objects[ColObjects,GS.Row]).getvalue('ACU_AFFAIRE');
V_PGI.DispatchTT( 5, taModif, CodeAffaire, '', 'MONOFICHE') ;
end;

procedure TOF_AFMODIFCUTOFF_MUL.GSFindLigne(Sender: TObject);
begin
if csDestroying in Ecran.ComponentState then Exit ;
Rechercher (GS, FindLigne, FindDebut) ;
end;

procedure TOF_AFMODIFCUTOFF_MUL.GSChercherClick(Sender: TObject);
begin
if csDestroying in Ecran.ComponentState then Exit ;
if GS.RowCount < 3 then Exit;
// PL le 29/11/02 : Ne pas forcer sur le grid : en cas de protection sur la première affaire bloquee,
// se positionne dessus alors qu'on ne devrait pas
//GS.SetFocus;
TPageControl(GetControl('PAGES')).ActivePage:=TTabSheet(getcontrol('PCRITERE'));
SetFocusControl('ACU_AFFAIRE1');
////////////////////////
FindDebut:=True ; FindLigne.Execute ;
end;

procedure TOF_AFMODIFCUTOFF_MUL.GSChargeTotaux(var CTOTAL: THNumEdit; CHTOTAL,CHCUMUL: string; ColTOTAL: integer);
var  Coord : TRect;
begin
  if (CTOTAL = nil) then
  begin
    CTOTAL := THNumEdit(GetControl(CHTOTAL));
    CTOTAL.ParentColor:=True;
    CTOTAL.Font.Style:=PCUMULSAI.Font.Style;
    CTOTAL.Font.Size:=PCUMULSAI.Font.Size;
    CTOTAL.Masks.PositiveMask:=GS.ColFormats[ColTOTAL];
    CTOTAL.Ctl3D:=False; CTOTAL.Top:=-1;
    Coord:=GS.CellRect(ColTOTAL,0);
    CTOTAL.Left:=Coord.Left + 1;
    CTOTAL.Width:=GS.ColWidths[ColTOTAL] + 1;
    CTOTAL.Height:=PCUMULSAI.Height;
    CTOTAL.Value:=TobAfcumul.Somme(CHCUMUL,[''],[''],False);
  end else
  begin
    Coord:=GS.CellRect(ColTOTAL,0);
    CTOTAL.Left:=Coord.Left + 1;
    CTOTAL.Width:=GS.ColWidths[ColTOTAL] + 1;
    CTOTAL.Value:=TobAfcumul.Somme(CHCUMUL,[''],[''],False);
  end;
end;

procedure TOF_AFMODIFCUTOFF_MUL.BEtatImprimerClick(Sender: TObject);
begin
  lanceEtattob ('E','ACF','ACG',TobAfcumul,true,false,false,nil,'','Edition Cut Off ',false);
end;

procedure TOF_AFMODIFCUTOFF_MUL.GSColumnWidthsChanged(Sender: TObject);
begin
  if PCUMULSAI.ControlCount <= 0 then exit ;
  GSChargeTotaux (TOTFAE,'TOTFAE','CUMFAE',ColFAE);
  GSChargeTotaux (TOTPCA,'TOTPCA','CUMPCA',ColPCA);
  GSChargeTotaux (TOTAAE,'TOTAAE','CUMAAE',ColAAE);
end;

/////////////// Procedure appellé par le bouton Validation //////////////
procedure AGLAfModifCutOff(parms:array of variant; nb: integer ) ;
var  F : TForm ;
     MaTOF  : TOF;
begin
F:=TForm(Longint(Parms[0])) ;
if (F is TFmul) then MaTOF:=TFMul(F).LaTOF else exit;
if (MaTOF is TOF_AFMODIFCUTOFF_MUL) then TOF_AFMODIFCUTOFF_MUL(MaTOF).AfModifCutOff else exit;
end;                            


Procedure AFLanceFiche_ModifCutOffMul(Argument:String);
begin
AGLLanceFiche ('AFF','AFMODIFCUTOFF_MUL','','',Argument);
end;

Initialization
registerclasses ( [ TOF_AFMODIFCUTOFF_MUL ] ) ;
RegisterAglProc('AfModifCutOff',TRUE,1,AGLAfModifCutOff);
end.
