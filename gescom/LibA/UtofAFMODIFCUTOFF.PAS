{***********UNITE*************************************************
Auteur  ...... : PL
Créé le ...... : 03/07/2002
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : AFMODIFCUTOFF ()
Mots clefs ... : TOF;AFMODIFCUTOFF
*****************************************************************}
Unit UtofAFMODIFCUTOFF;

Interface

Uses StdCtrls,
     Controls,  
     Classes,
     windows,
     messages,
     ent1,
     HTB97,
{$IFDEF EAGLCLIENT}
Maineagl,
{$ELSE}
       db,  {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}FE_Main,
{$ENDIF}
     forms,M3FP,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HPanel,
     HMsgBox,
     HSysMenu,
     paramsoc,
     UTOF, utob, vierge,
     EntGC, AfUtilArticle, AffaireUtil,
     SaisUtil, dicobtp, FactUtil, utilCutoff, ActiviteUtil, FactArticle, 
     uafo_ressource,
     UTofAfBaseCodeAffaire,
     UtofAFMODIFCUTOFFADD ;

Type
  TOF_AFMODIFCUTOFF = Class (TOF_AFBASECODEAFFAIRE)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnClose                  ; override ;
    procedure bFermerOnClick(SEnder: TObject);
    private
        GS : THGrid;
        LesCol : String;
        nbCol, ColAff,ColArt,ColLibArt,ColRess,ColLibRes,ColTypA,ColLibTypA,ColProd,ColFact,ColSolde,ColVide,ColFAE,ColAAE,ColPCA : integer;
        TobReg, TobAchat : TOB;
        bModified, bTotModified, bTotAchatModified : boolean;
        procedure DessineTotaux;
        procedure GSCellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
        procedure FormResize(Sender: TObject);
        procedure AfficheGrid;
        procedure TOTChange(Sender: TObject);
        procedure TOTAchatChange(Sender: TObject);
        procedure AfficheLibellesOnClick(Sender: TObject);
        procedure AfficheLignesVidesOnClick(Sender: TObject);
    public
        Action   :  TActionFiche ;
        EnErreur :  boolean;
        GModeEclat: T_ModeEclat;
        AFOAssistants : TAFO_Ressources;
        TOBArticles : TOB;
        TOBAssistant, TOBArt : TOB;
        procedure NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2, Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers, Tiers_:THEdit); override;
        procedure Eclatement(AcbParle:boolean);
        procedure Insertion;
        procedure ClearLesMontantsSituation (TT:TOB);
        procedure ReinitLesTotaux;

  end ;

function AFLanceFiche_ModifCutOff(Argument:String):string;

Implementation

procedure TOF_AFMODIFCUTOFF.OnArgument (S : String ) ;
Var
    Critere, Champ, valeur, St  : String;
    x : integer;
begin
TobAchat := nil;
Action:=taModif ;
bModified:=false;
bTotModified:=false;
bTotAchatModified := false;
// Recup des critères
Critere:=(Trim(ReadTokenSt(S)));
While (Critere <>'') do
    BEGIN
    if Critere<>'' then
        BEGIN
        x:=pos('=',Critere);
        if x<>0 then
           begin
           Champ:=copy(Critere,1,X-1);
           Valeur:=Copy (Critere,X+1,length(Critere)-X);
           end;

        if Champ='ACTION' then
             begin
             if Valeur='CREATION' then BEGIN Action:=taCreat ; END ;
             if Valeur='MODIFICATION' then BEGIN Action:=taModif ; END ;
             if Valeur='CONSULTATION' then BEGIN Action:=taConsult ; END ;
             end;
        if Champ = 'TIERS' then SetControlText('ACU_TIERS',Valeur);
        if Champ = 'AFFAIRE' then SetControlText('ACU_AFFAIRE',Valeur);
        if Champ = 'DATE' then SetControlText('ACU_DATE',Valeur);
        if Champ = 'MOISCLOTURE' then SetControlText('MOISCLOTURE',Valeur);
        if Champ = 'BLOCAGE' then begin SetControlVisible('TAFFAIREBLOQUEE', true); Action:=taConsult; end ;
        END;
    Critere:=(Trim(ReadTokenSt(S)));
    END;

if Action=taConsult then
  BEGIN
  FicheReadOnly(Ecran) ;
  SetControlEnabled('BECLATER', false);
  END;

  If not(ctxScot in V_PGI.PGIContexte)  then
  begin
      SetControlVisible ('TMOISCLOTURE' , False);
      SetControlVisible ('MOISCLOTURE' , False);
  end;


// Init du code affaire dans la tof ancêtre
Inherited ;
// Force la protection des champs Affaire
EditAff.ReadOnly := true;
EditAff1.ReadOnly := true;
EditAff2.ReadOnly := true;
EditAff3.ReadOnly := true;
EditAff4.ReadOnly := true;

AFOAssistants:=TAFO_Ressources.Create;
TOBAssistant:=TOB.Create('RESSOURCE',Nil,-1) ;
TOBArt := Nil;
TOBArticles:=TOB.Create('Les Articles',Nil,-1) ;

// Mode d'eclatement
GModeEclat:=DetermineModeEclatCutOff;
// Gestion du Grid
GS := THGRID(GetControl('GS'));
case GModeEclat of
   tmeRessArt :begin
                LesCol := 'ACU_AFFAIRE;ACU_RESSOURCE;LIBRESSOURCE;ACU_CODEARTICLE;LIBARTICLE';
                nbCol := 12;
                end;
   tmeArticle :begin
                LesCol := 'ACU_AFFAIRE;ACU_CODEARTICLE;LIBARTICLE';
                nbCol := 10;
                end;
   tmeRessource :begin
                LesCol := 'ACU_AFFAIRE;ACU_RESSOURCE;LIBRESSOURCE';
                nbCol := 10;
                end;
   tmeRessTypA :begin
                LesCol := 'ACU_AFFAIRE;ACU_RESSOURCE;LIBRESSOURCE;ACU_TYPEARTICLE;LIBTYPEARTICLE';
                nbCol := 12;
                end;
   tmeGlobal :begin
                LesCol := 'ACU_AFFAIRE;';
                nbCol := 8;
                end;
   tmeSans :begin
                LesCol := 'ACU_AFFAIRE;';
                nbCol := 8;
                end;
   tmeTypeArticle :begin
                LesCol := 'ACU_AFFAIRE;ACU_TYPEARTICLE;LIBTYPEARTICLE';
                nbCol := 10;
                end;
end;
LesCol := LesCol + ';ACU_PVPROD;ACU_PVFACT;ACU_CUTOFFORIG;VIDE;ACU_FAE;ACU_AAE;ACU_PCA';
GS.ColCount:=NbCol;

St:=LesCol;
for x:=0 to GS.ColCount-1 do
   BEGIN
   if x>2 then  GS.ColWidths[x]:=100;
   Champ:=ReadTokenSt(St) ;
   if Champ='ACU_RESSOURCE' then ColRess := x
   else if Champ='LIBRESSOURCE' then ColLibRes := x
   else if Champ='LIBARTICLE' then ColLibArt := x
   else if Champ='LIBTYPEARTICLE' then ColLibTypA := x
   else if Champ='ACU_AFFAIRE' then ColAff := x
   else if Champ='ACU_CODEARTICLE' then ColArt := x
   else if Champ='ACU_TYPEARTICLE' then ColTypA := x
   else if Champ='ACU_PVPROD' then ColProd := x
   else if Champ='ACU_PVFACT' then ColFact := x
   else if Champ='ACU_CUTOFFORIG' then ColSolde := x
   else if Champ='VIDE'  then ColVide:= x
   else if Champ='ACU_FAE'   then ColFAE := x
   else if Champ='ACU_AAE'   then ColAAE := x
   else if Champ='ACU_PCA'   then ColPCA:= x
   ;
   END ;

// dessin des deux ou trois premières colonnes
GS.ColWidths[ColAff]:=0;
GS.ColLengths[ColAff]:=-1;
case GModeEclat of
   tmeRessArt :begin
                GS.Cells[ColRess,0]:=TraduitGA('Ressource');
                GS.Cells[ColLibRes,0]:=TraduitGA('Libellé Ressource');
                GS.Cells[ColArt,0]:=TraduitGA('Article');
                GS.Cells[ColLibArt,0]:=TraduitGA('Libellé Article');
                GS.ColWidths[ColRess]:=0;
                GS.ColWidths[ColLibRes]:=200;
                GS.ColWidths[ColArt]:=0;
                GS.ColWidths[ColLibArt]:=100;
                GS.ColLengths[ColRess]:=-1;
                GS.ColLengths[ColLibRes]:=-1;
                GS.ColLengths[ColArt]:=-1;
                GS.ColLengths[ColLibArt]:=-1;
                end;
   tmeArticle :begin
                GS.Cells[ColArt,0]:=TraduitGA('Article');
                GS.Cells[ColLibArt,0]:=TraduitGA('Libellé Article');
                GS.ColWidths[ColArt]:=0;
                GS.ColWidths[ColLibArt]:=200;
                GS.ColLengths[ColArt]:=-1;
                GS.ColLengths[ColLibArt]:=-1;
                end;
   tmeRessource :begin
                GS.Cells[ColRess,0]:=TraduitGA('Ressource');
                GS.Cells[ColLibRes,0]:=TraduitGA('Libellé Ressource');
                GS.ColWidths[ColRess]:=0;
                GS.ColWidths[ColLibRes]:=300;
                GS.ColLengths[ColRess]:=-1;
                GS.ColLengths[ColLibRes]:=-1;
                end;
   tmeRessTypA :begin
                GS.Cells[ColRess,0]:=TraduitGA('Ressource');
                GS.Cells[ColLibRes,0]:=TraduitGA('Libellé Ressource');
                GS.Cells[ColTypA,0]:=TraduitGA('Type');
                GS.Cells[ColLibTypA,0]:=TraduitGA('Libellé Type');
                GS.ColWidths[ColRess]:=0;
                GS.ColWidths[ColLibRes]:=200;
                GS.ColWidths[ColTypA]:=0;
                GS.ColWidths[ColLibTypA]:=100;
                GS.ColLengths[ColRess]:=-1;
                GS.ColLengths[ColTypA]:=-1;
                GS.ColLengths[ColLibTypA]:=-1;
                GS.ColLengths[ColLibRes]:=-1;
                // PL le 20/02/03 : Spécif Amyot, l'écran est différent
                if (TForm(Ecran).Name = 'AFMODIFCUTOFF') then
                  begin
                  // PL le 07/01/03 : Gestion Amyot et standard : on gère les types articles en éclatés
                  SetControlVisible('TLBLPrestation', true);
                  SetControlVisible('TLBLFrais', true);
                  SetControlVisible('TLBLFourniture', true);
                  SetControlVisible('TOTSOLDE0', true);
                  SetControlVisible('TOTSOLDE1', true);
                  SetControlVisible('TOTSOLDE2', true);
                  SetControlVisible('TOTFAE0', true);
                  SetControlVisible('TOTFAE1', true);
                  SetControlVisible('TOTFAE2', true);
                  SetControlVisible('TOTAAE0', true);
                  SetControlVisible('TOTAAE1', true);
                  SetControlVisible('TOTAAE2', true);
                  SetControlVisible('TOTPCA0', true);
                  SetControlVisible('TOTPCA1', true);
                  SetControlVisible('TOTPCA2', true);
                  SetControlVisible('LblSaisieEclat', true);
                  SetControlVisible('PanelLigne', true);
                  SetControlProperty('TOTFAE', 'ReadOnly', true);
                  SetControlProperty('TOTFAE', 'tabstop', false);
                  SetControlProperty('TOTAAE', 'ReadOnly', true);
                  SetControlProperty('TOTAAE', 'tabstop', false);
                  SetControlProperty('TOTPCA', 'ReadOnly', true);
                  SetControlProperty('TOTPCA', 'tabstop', false);
                  end;
                ////////////////////////////
                end;
   tmeGlobal :begin
                GS.Cells[ColAff,0]:=TraduitGA('Affaire');
                GS.ColWidths[ColAff]:=200;
                end;
   tmeSans :begin
                GS.Cells[ColAff,0]:=TraduitGA('Affaire');
                GS.ColWidths[ColAff]:=200;
                end;
   tmeTypeArticle :begin
                GS.Cells[ColTypA,0]:=TraduitGA('Type');
                GS.Cells[ColLibTypA,0]:=TraduitGA('Libellé Type');
                GS.ColWidths[ColTypA]:=0;
                GS.ColWidths[ColLibTypA]:=200;
                GS.ColLengths[ColTypA]:=-1;
                GS.ColLengths[ColLibTypA]:=-1;
                end;
end;

// largeurs des colonnes
// PL le 25/10/02 : problème pour le calcul de la colonne FACTURE, on la fait disparaitre
GS.ColWidths[ColFact]:=0;
// PL le 07/01/03 : Embiguité amyot pour la colonne production, on la fait disparaitre
GS.ColWidths[ColProd]:=0;
//GS.ColWidths[ColProd]:=100;
//GS.ColWidths[ColFact]:=100;
////////////// fin modif PL le 25/10/02
GS.ColWidths[ColSolde]:=110;
GS.ColWidths[ColVide]:=10;
GS.ColWidths[ColFAE]:=100;
GS.ColWidths[ColAAE]:=100;
GS.ColWidths[ColPCA]:=100;


// Protection des colonnes
GS.ColLengths[ColProd]:=-1;
GS.ColLengths[ColFact]:=-1;
GS.ColLengths[ColSolde]:=-1;
GS.ColLengths[ColVide]:=-1;

// justif des colonnes
GS.ColAligns[ColProd]:=taRightJustify;
GS.ColAligns[ColFact]:=taRightJustify;
GS.ColAligns[ColSolde]:=taRightJustify;
GS.ColAligns[ColFAE]:=taRightJustify;
GS.ColAligns[ColAAE]:=taRightJustify;
GS.ColAligns[ColPCA]:=taRightJustify;

// libellés des colonnes
GS.Cells[ColProd,0]:= TraduitGA('Production');
GS.Cells[ColFact,0]:= TraduitGA('Facturé');
GS.Cells[ColSolde,0]:= TraduitGA('Résultat formule');
GS.Cells[ColVide,0]:= '';
GS.Cells[ColFAE,0]:= TraduitGA('FAE');
GS.Cells[ColAAE,0]:= TraduitGA('AAE');
GS.Cells[ColPCA,0]:= TraduitGA('PCA');


// dessin des zones totaux aux bons emplacements sous les colonnes du grid associees
DessineTotaux;

// Chargement de la TOB pour affichage du GRID des cut off
AfficheGrid;

GS.OnCellExit := GSCellExit;
TForm(Ecran).OnResize:=FormResize;

if (TForm(Ecran).Name = 'AFMODIFCUTOFF') then
  begin
  if (GModeEclat = tmeRessTypA) then
    begin
    TForm(Ecran).ActiveControl := THNumEdit(GetControl('TOTFAE0'));
    ChangeMask(THNumEdit(GetControl('TOTFAE0')), V_PGI.OkDecV, '') ;
    ChangeMask(THNumEdit(GetControl('TOTAAE0')), V_PGI.OkDecV, '') ;
    ChangeMask(THNumEdit(GetControl('TOTPCA0')), V_PGI.OkDecV, '') ;
    THNumEdit(GetControl('TOTFAE0')).onchange := TOTChange;
    THNumEdit(GetControl('TOTAAE0')).onchange := TOTChange;
    THNumEdit(GetControl('TOTPCA0')).onchange := TOTChange;
    ChangeMask(THNumEdit(GetControl('TOTFAE1')), V_PGI.OkDecV, '') ;
    ChangeMask(THNumEdit(GetControl('TOTAAE1')), V_PGI.OkDecV, '') ;
    ChangeMask(THNumEdit(GetControl('TOTPCA1')), V_PGI.OkDecV, '') ;
    THNumEdit(GetControl('TOTFAE1')).onchange := TOTChange;
    THNumEdit(GetControl('TOTAAE1')).onchange := TOTChange;
    THNumEdit(GetControl('TOTPCA1')).onchange := TOTChange;
    ChangeMask(THNumEdit(GetControl('TOTFAE2')), V_PGI.OkDecV, '') ;
    ChangeMask(THNumEdit(GetControl('TOTAAE2')), V_PGI.OkDecV, '') ;
    ChangeMask(THNumEdit(GetControl('TOTPCA2')), V_PGI.OkDecV, '') ;
    THNumEdit(GetControl('TOTFAE2')).onchange := TOTChange;
    THNumEdit(GetControl('TOTAAE2')).onchange := TOTChange;
    THNumEdit(GetControl('TOTPCA2')).onchange := TOTChange;
    end
  else
    begin
    ChangeMask(THNumEdit(GetControl('TOTFAE')), V_PGI.OkDecV, '') ;
    ChangeMask(THNumEdit(GetControl('TOTAAE')), V_PGI.OkDecV, '') ;
    ChangeMask(THNumEdit(GetControl('TOTPCA')), V_PGI.OkDecV, '') ;
    THNumEdit(GetControl('TOTFAE')).onchange := TOTChange;
    THNumEdit(GetControl('TOTAAE')).onchange := TOTChange;
    THNumEdit(GetControl('TOTPCA')).onchange := TOTChange;
    end;
  end
else
// PL le 20/02/03 : Spécif Amyot, l'écran est différent
if (TForm(Ecran).Name = 'AFMODIFCUTOFFAMYO') then
  begin
    TForm(Ecran).ActiveControl := THNumEdit(GetControl('TOTFAR'));
    ChangeMask(THNumEdit(GetControl('TOTFAR')), V_PGI.OkDecV, '') ;
    ChangeMask(THNumEdit(GetControl('TOTAAR')), V_PGI.OkDecV, '') ;
    ChangeMask(THNumEdit(GetControl('TOTCCA')), V_PGI.OkDecV, '') ;
    THNumEdit(GetControl('TOTFAR')).onchange := TOTAchatChange;
    THNumEdit(GetControl('TOTAAR')).onchange := TOTAchatChange;
    THNumEdit(GetControl('TOTCCA')).onchange := TOTAchatChange;
    ChangeMask(THNumEdit(GetControl('TOTFAE1')), V_PGI.OkDecV, '') ;
    ChangeMask(THNumEdit(GetControl('TOTAAE1')), V_PGI.OkDecV, '') ;
    ChangeMask(THNumEdit(GetControl('TOTPCA1')), V_PGI.OkDecV, '') ;
    THNumEdit(GetControl('TOTFAE1')).onchange := TOTChange;
    THNumEdit(GetControl('TOTAAE1')).onchange := TOTChange;
    THNumEdit(GetControl('TOTPCA1')).onchange := TOTChange;
    ChangeMask(THNumEdit(GetControl('TOTFAEACHAT')), V_PGI.OkDecV, '') ;
    ChangeMask(THNumEdit(GetControl('TOTAAEACHAT')), V_PGI.OkDecV, '') ;
    ChangeMask(THNumEdit(GetControl('TOTPCAACHAT')), V_PGI.OkDecV, '') ;
    THNumEdit(GetControl('TOTFAEACHAT')).onchange := TOTChange;
    THNumEdit(GetControl('TOTAAEACHAT')).onchange := TOTChange;
    THNumEdit(GetControl('TOTPCAACHAT')).onchange := TOTChange;
  end;

// Colonne courante dans le grid
GS.Col:= ColFAE;
// Pour recadrage à doite correct dans le grid (mauvais fonctionnement du resize)
GS.ColWidths[ColVide] := GS.ColWidths[ColVide] - 3;

TToolBarButton97(GetControl('BFERME')).onclick := bFermerOnClick;
TCheckBox(GetControl('AFFICHELIBELLES')).onclick := AfficheLibellesOnClick;
// PL le 19/05/03 : Spécif Amyot, l'écran est différent
if (TForm(Ecran).Name = 'AFMODIFCUTOFFAMYO') then
  begin
    TCheckBox(GetControl('VOIRLIGNESVIDES')).onclick := AfficheLignesVidesOnClick;
  end;
end ;

procedure TOF_AFMODIFCUTOFF.TOTAchatChange (Sender : TObject);
  var
  SommePrestation, SommeFournitures, SommeFAR, SommeAAR, SommeCCA : double;
begin
  bTotAchatModified := true;

  THNumEdit(Sender).value := abs (THNumEdit(Sender).value);

  SommePrestation := THNumEdit(GetControl ('TOTSOLDE0')).Value;
  SommeFournitures := THNumEdit(GetControl ('TOTSOLDE2')).Value;
  SommeFAR := THNumEdit(GetControl ('TOTFAR')).Value;
  SommeAAR := THNumEdit(GetControl ('TOTAAR')).Value;
  SommeCCA := THNumEdit(GetControl ('TOTCCA')).Value;

  // la base de ventilation Cut off Amyot = Somme Prestation + Somme Fourniture + Somme des FAR - Somme des AAR - Somme des CCA
  THNumEdit(GetControl ('BASEVENTIL')).Value := SommePrestation + SommeFournitures + SommeFAR - SommeAAR - SommeCCA;
end;

procedure TOF_AFMODIFCUTOFF.AfficheLignesVidesOnClick (Sender: TObject);
  var
  ii, jj : integer;
  bTotModifMem : boolean;
begin
  bTotModifMem := bTotAchatModified;
  TFVierge(Ecran).BValiderClick (nil);

  try
  GS.CacheEdit;
  GS.SynEnabled := False;

  ii := GS.RowCount - 1;
  for jj := 1 to ii do
    begin
      GS.DeleteRow (1);
    end;

  if GS.RowCount > 1 then
    GS.Row := 1;

  TobReg.Free;
  TobReg := nil;

  AfficheGrid;

  finally
    GS.SynEnabled := true;
    GS.MontreEdit;
    bTotAchatModified := bTotModifMem;
  end;

end;

procedure TOF_AFMODIFCUTOFF.AfficheLibellesOnClick (Sender: TObject);
begin
If (GetCheckBoxState('AFFICHELIBELLES') = cbChecked) then
  begin
    case GModeEclat of
      tmeRessArt :begin
                GS.ColWidths[ColRess]:=0;
                GS.ColWidths[ColLibRes]:=200;
                GS.ColWidths[ColArt]:=0;
                GS.ColWidths[ColLibArt]:=100;
                end;
      tmeArticle :begin
                GS.ColWidths[ColArt]:=0;
                GS.ColWidths[ColLibArt]:=200;
                end;
      tmeRessource :begin
                GS.ColWidths[ColRess]:=0;
                GS.ColWidths[ColLibRes]:=300;
                end;
      tmeRessTypA :begin
                GS.ColWidths[ColRess]:=0;
                GS.ColWidths[ColLibRes]:=200;
                GS.ColWidths[ColTypA]:=0;
                GS.ColWidths[ColLibTypA]:=100;
                end;
      tmeGlobal :begin
                GS.ColWidths[ColAff]:=200;
                end;
      tmeSans :begin
                GS.ColWidths[ColAff]:=200;
                end;
      tmeTypeArticle :begin
                GS.ColWidths[ColTypA]:=0;
                GS.ColWidths[ColLibTypA]:=200;
                end;
    end;
  end
else
  begin
    case GModeEclat of
      tmeRessArt :begin
                GS.ColWidths[ColRess]:=200;
                GS.ColWidths[ColLibRes]:=0;
                GS.ColWidths[ColArt]:=100;
                GS.ColWidths[ColLibArt]:=0;
                end;
      tmeArticle :begin
                GS.ColWidths[ColArt]:=200;
                GS.ColWidths[ColLibArt]:=0;
                end;
      tmeRessource :begin
                GS.ColWidths[ColRess]:=300;
                GS.ColWidths[ColLibRes]:=0;
                end;
      tmeRessTypA :begin
                GS.ColWidths[ColRess]:=200;
                GS.ColWidths[ColLibRes]:=0;
                GS.ColWidths[ColTypA]:=100;
                GS.ColWidths[ColLibTypA]:=0;
                end;
      tmeGlobal :begin
                GS.ColWidths[ColAff]:=200;
                end;
      tmeSans :begin
                GS.ColWidths[ColAff]:=200;
                end;
      tmeTypeArticle :begin
                GS.ColWidths[ColTypA]:=200;
                GS.ColWidths[ColLibTypA]:=0;
                end;
    end;
  end;

DessineTotaux;
end;

procedure TOF_AFMODIFCUTOFF.TOTChange(Sender: TObject);
begin
  THNumEdit (Sender).value := abs (THNumEdit(Sender).value);

  if (GModeEclat = tmeRessTypA) then
    begin
    if (Ecran.activecontrol=GetControl('TOTFAE0')) or (Ecran.activecontrol=GetControl('TOTAAE0')) or (Ecran.activecontrol=GetControl('TOTPCA0'))
      or (Ecran.activecontrol=GetControl('TOTFAE1')) or (Ecran.activecontrol=GetControl('TOTAAE1')) or (Ecran.activecontrol=GetControl('TOTPCA1'))
      or (Ecran.activecontrol=GetControl('TOTFAE2')) or (Ecran.activecontrol=GetControl('TOTAAE2')) or (Ecran.activecontrol=GetControl('TOTPCA2')) then
        bTotModified:=true;
    end
  else
    if (Ecran.activecontrol=GetControl('TOTFAE')) or (Ecran.activecontrol=GetControl('TOTAAE')) or (Ecran.activecontrol=GetControl('TOTPCA')) then
      bTotModified:=true;

end;


procedure TOF_AFMODIFCUTOFF.GSCellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
var
St, StC : String ;
dSommeFAE, dSommeAAE, dSommePCA : double;
dSommeFAE0, dSommeAAE0, dSommePCA0 : double;
dSommeFAE1, dSommeAAE1, dSommePCA1 : double;
dSommeFAE2, dSommeAAE2, dSommePCA2 : double;
begin
if Action=taConsult then Exit;
dSommeFAE0 := 0; dSommeAAE0 := 0; dSommePCA0 := 0; dSommeFAE1 := 0; dSommeAAE1 := 0; dSommePCA1 := 0;
dSommeFAE2 := 0; dSommeAAE2 := 0; dSommePCA2 := 0;

//
// Gestion du format de saisie des cellules
//
St:=GS.Cells[ACol,ARow];
// PL le 24/10/02 : modif de dernière minute suite remarques AMYOT : ajout de la valeur absolue abs() sur la valeur
if (St = '') then St := '0';
if (ACol in [ColFAE, ColAAE, ColPCA]) then StC := StrF00 (abs(Valeur(St)), V_PGI.OkDecV) else
   StC := St;

if (StC = '') then StC := '0';
/////////// fin modif PL le 24/10/02


GS.Cells[ACol,ARow] := StC;

//
// Gestion des sommes des colonnes du grid
//
// Récupération des données saisies dans le grid
TobReg.GetGridDetail (GS, GS.RowCount-1, 'AFCUMUL', LesCol);
dSommeFAE := TobReg.Somme('ACU_FAE', ['ACU_AFFAIRE'], [GetControlText('ACU_AFFAIRE')], true);
dSommeAAE := TobReg.Somme('ACU_AAE', ['ACU_AFFAIRE'], [GetControlText('ACU_AFFAIRE')], true);
dSommePCA := TobReg.Somme('ACU_PCA', ['ACU_AFFAIRE'], [GetControlText('ACU_AFFAIRE')], true);
if (GModeEclat = tmeRessTypA) then
  begin
  dSommeFAE0 := TobReg.Somme('ACU_FAE', ['ACU_AFFAIRE','ACU_TYPEARTICLE'], [GetControlText('ACU_AFFAIRE'),'PRE'], true);
  dSommeAAE0 := TobReg.Somme('ACU_AAE', ['ACU_AFFAIRE','ACU_TYPEARTICLE'], [GetControlText('ACU_AFFAIRE'),'PRE'], true);
  dSommePCA0 := TobReg.Somme('ACU_PCA', ['ACU_AFFAIRE','ACU_TYPEARTICLE'], [GetControlText('ACU_AFFAIRE'),'PRE'], true);
  dSommeFAE1 := TobReg.Somme('ACU_FAE', ['ACU_AFFAIRE','ACU_TYPEARTICLE'], [GetControlText('ACU_AFFAIRE'),'FRA'], true);
  dSommeAAE1 := TobReg.Somme('ACU_AAE', ['ACU_AFFAIRE','ACU_TYPEARTICLE'], [GetControlText('ACU_AFFAIRE'),'FRA'], true);
  dSommePCA1 := TobReg.Somme('ACU_PCA', ['ACU_AFFAIRE','ACU_TYPEARTICLE'], [GetControlText('ACU_AFFAIRE'),'FRA'], true);
  dSommeFAE2 := TobReg.Somme('ACU_FAE', ['ACU_AFFAIRE','ACU_TYPEARTICLE'], [GetControlText('ACU_AFFAIRE'),'MAR'], true);
  dSommeAAE2 := TobReg.Somme('ACU_AAE', ['ACU_AFFAIRE','ACU_TYPEARTICLE'], [GetControlText('ACU_AFFAIRE'),'MAR'], true);
  dSommePCA2 := TobReg.Somme('ACU_PCA', ['ACU_AFFAIRE','ACU_TYPEARTICLE'], [GetControlText('ACU_AFFAIRE'),'MAR'], true);
  end;

// Affectation des FAE, AAE, PCA
if (TobReg<>nil) then
    begin
    TobReg.Detail[ARow-1].PutValue('ACU_FAEDEV', TobReg.Detail[ARow-1].GetValue('ACU_FAE'));
    TobReg.Detail[ARow-1].PutValue('ACU_AAEDEV', TobReg.Detail[ARow-1].GetValue('ACU_AAE'));
    TobReg.Detail[ARow-1].PutValue('ACU_PCADEV', TobReg.Detail[ARow-1].GetValue('ACU_PCA'));
    end;

if (GModeEclat = tmeRessTypA) then
  begin
  if (TFVierge(Ecran).Name = 'AFMODIFCUTOFF') then
    begin
    if (THNumEdit(GetControl('TOTFAE0')).Value <> dSommeFAE0)
      or (THNumEdit(GetControl('TOTAAE0')).Value <> dSommeAAE0)
      or (THNumEdit(GetControl('TOTPCA0')).Value <> dSommePCA0) then
      begin
      bModified:=true;
      TobReg.Detail[ARow-1].Putvalue('ACU_DATEMODIF', V_PGI.DateEntree);
      end;
    if (THNumEdit(GetControl('TOTFAE1')).Value <> dSommeFAE1)
      or (THNumEdit(GetControl('TOTAAE1')).Value <> dSommeAAE1)
      or (THNumEdit(GetControl('TOTPCA1')).Value <> dSommePCA1) then
      begin
      bModified:=true;
      TobReg.Detail[ARow-1].Putvalue('ACU_DATEMODIF', V_PGI.DateEntree);
      end;
    if (THNumEdit(GetControl('TOTFAE2')).Value <> dSommeFAE2)
      or (THNumEdit(GetControl('TOTAAE2')).Value <> dSommeAAE2)
      or (THNumEdit(GetControl('TOTPCA2')).Value <> dSommePCA2) then
      begin
      bModified:=true;
      TobReg.Detail[ARow-1].Putvalue('ACU_DATEMODIF', V_PGI.DateEntree);
      end;
    end
  else
  if (TFVierge(Ecran).Name = 'AFMODIFCUTOFFAMYO') then
    begin
    if (THNumEdit(GetControl('TOTFAEACHAT')).Value <> (dSommeFAE0 + dSommeFAE2))
      or (THNumEdit(GetControl('TOTAAEACHAT')).Value <> (dSommeAAE0 + dSommeAAE2))
      or (THNumEdit(GetControl('TOTPCAACHAT')).Value <> (dSommePCA0 + dSommePCA2)) then
      begin
      bModified:=true;
      TobReg.Detail[ARow-1].Putvalue('ACU_DATEMODIF', V_PGI.DateEntree);
      end;
    if (THNumEdit(GetControl('TOTFAE1')).Value <> dSommeFAE1)
      or (THNumEdit(GetControl('TOTAAE1')).Value <> dSommeAAE1)
      or (THNumEdit(GetControl('TOTPCA1')).Value <> dSommePCA1) then
      begin
      bModified:=true;
      TobReg.Detail[ARow-1].Putvalue('ACU_DATEMODIF', V_PGI.DateEntree);
      end;
    end;
  end
else
  begin
  if (THNumEdit(GetControl('TOTFAE')).Value <> dSommeFAE)
    or (THNumEdit(GetControl('TOTAAE')).Value <> dSommeAAE)
    or (THNumEdit(GetControl('TOTPCA')).Value <> dSommePCA) then
    begin
    bModified:=true;
    TobReg.Detail[ARow-1].Putvalue('ACU_DATEMODIF', V_PGI.DateEntree);
    end;
  end;

THNumEdit(GetControl('TOTFAE')).Value := Abs(dSommeFAE);
THNumEdit(GetControl('TOTAAE')).Value := Abs(dSommeAAE);
THNumEdit(GetControl('TOTPCA')).Value := Abs(dSommePCA);
if (GModeEclat = tmeRessTypA) then
  begin
  if (TFVierge(Ecran).Name = 'AFMODIFCUTOFF') then
    begin
    THNumEdit(GetControl('TOTFAE0')).Value := Abs(dSommeFAE0);
    THNumEdit(GetControl('TOTAAE0')).Value := Abs(dSommeAAE0);
    THNumEdit(GetControl('TOTPCA0')).Value := Abs(dSommePCA0);
    THNumEdit(GetControl('TOTFAE1')).Value := Abs(dSommeFAE1);
    THNumEdit(GetControl('TOTAAE1')).Value := Abs(dSommeAAE1);
    THNumEdit(GetControl('TOTPCA1')).Value := Abs(dSommePCA1);
    THNumEdit(GetControl('TOTFAE2')).Value := Abs(dSommeFAE2);
    THNumEdit(GetControl('TOTAAE2')).Value := Abs(dSommeAAE2);
    THNumEdit(GetControl('TOTPCA2')).Value := Abs(dSommePCA2);
    end
  else
  if (TFVierge(Ecran).Name = 'AFMODIFCUTOFFAMYO') then
    begin
    THNumEdit(GetControl('TOTFAEACHAT')).Value := Abs(dSommeFAE0 + dSommeFAE2);
    THNumEdit(GetControl('TOTAAEACHAT')).Value := Abs(dSommeAAE0 + dSommeAAE2);
    THNumEdit(GetControl('TOTPCAACHAT')).Value := Abs(dSommePCA0 + dSommePCA2);
    THNumEdit(GetControl('TOTFAE1')).Value := Abs(dSommeFAE1);
    THNumEdit(GetControl('TOTAAE1')).Value := Abs(dSommeAAE1);
    THNumEdit(GetControl('TOTPCA1')).Value := Abs(dSommePCA1);
    end;
  end;
end;

procedure TOF_AFMODIFCUTOFF.FormResize(Sender: TObject);
begin
DessineTotaux;
end;

procedure TOF_AFMODIFCUTOFF.AfficheGrid;
var
   TobDet : TOB;
   QQ : Tquery;
   Req, ReqAchat: String;
   wi, iIndiceRess : Integer;
   dSommeSolde, dSommeFAE, dSommeAAE, dSommePCA : double;
   dSommeSolde0, dSommeFAE0, dSommeAAE0, dSommePCA0 : double;
   dSommeSolde1, dSommeFAE1, dSommeAAE1, dSommePCA1 : double;
   dSommeSolde2, dSommeFAE2, dSommeAAE2, dSommePCA2 : double;
   dSommeFAR, dSommeAAR, dSommeCCA : double;
   RechArt : T_RechArt ;
   bAchatOK : boolean;
begin
  bAchatOK := false;
  dSommeSolde:=0;  dSommeSolde0:=0;  dSommeSolde1:=0;  dSommeSolde2:=0;
  dSommeFAE:=0;  dSommeFAE0:=0;  dSommeFAE1:=0;  dSommeFAE2:=0;
  dSommeAAE:=0;  dSommeAAE0:=0;  dSommeAAE1:=0;  dSommeAAE2:=0;
  dSommePCA:=0;  dSommePCA0:=0;  dSommePCA1:=0;  dSommePCA2:=0;
  dSommeFAR:=0;  dSommeAAR:=0;  dSommeCCA:=0;
  TobReg := Tob.Create('Liste Cutt Off',nil,-1);
  // SELECT * : nombre de lignes restreint
  Req := 'SELECT * FROM AFCUMUL Where ACU_TYPEAC="CVE"'
          + ' AND ACU_DATE="' + usdatetime (strtodate (GetControlText ('ACU_DATE')))
          + '" AND ACU_TIERS="' + GetControlText ('ACU_TIERS')
          + '" AND ACU_AFFAIRE="' + GetControlText ('ACU_AFFAIRE') + '" ';

  // PL le 19/05/03 : spécid Amyot
  if (TForm(Ecran).Name = 'AFMODIFCUTOFFAMYO') then
    if (GetCheckBoxState('VOIRLIGNESVIDES') = cbUnchecked) then
      Req := Req + 'AND (ACU_AAE<>0 OR ACU_FAE<>0 OR ACU_PCA<>0) ';

  case GModeEclat of
     tmeRessArt :begin
                  Req := Req + 'ORDER BY ACU_RESSOURCE,ACU_CODEARTICLE';
                  end;
     tmeArticle :begin
                  Req := Req + 'ORDER BY ACU_CODEARTICLE';
                  end;
     tmeRessource :begin
                  Req := Req + 'ORDER BY ACU_RESSOURCE';
                  end;
     tmeRessTypA :begin
                  Req := Req + 'ORDER BY ACU_RESSOURCE,ACU_TYPEARTICLE';
                  end;
     tmeGlobal :begin
                  Req := Req + 'ORDER BY ACU_AFFAIRE';
                  end;
     tmeSans :begin
                  Req := Req + 'ORDER BY ACU_AFFAIRE';
                  end;
     tmeTypeArticle :begin
                  Req := Req + 'ORDER BY ACU_TYPEARTICLE';
                  end;
  end;

  QQ := nil;
  try
    QQ := OpenSQL (Req, True);
    If Not QQ.EOF then TobReg.LoadDetailDB ('AFCUMUL', '', '', QQ, True)
      Else Exit;

  Finally
    Ferme (QQ);
    if (TobReg.Detail.count <> 0) then
      GS.RowCount := TobReg.Detail.count + 1;
  End;

  if (TobReg.Detail.Count > 0) Then
    for wi:=0 To TobReg.Detail.Count-1 do
      Begin
      TobDet := TobReg.Detail[wi];
      TobDet.AddChampSup ('LIBRESSOURCE', false);
      TobDet.AddChampSup ('LIBARTICLE', false);
      TobDet.AddChampSup ('LIBTYPEARTICLE', false);
      dSommeSolde := dSommeSolde + TobDet.GetValue('ACU_CUTOFFORIG');
      dSommeFAE := dSommeFAE + TobDet.GetValue('ACU_FAE');
      dSommeAAE := dSommeAAE + TobDet.GetValue('ACU_AAE');
      dSommePCA := dSommePCA + TobDet.GetValue('ACU_PCA');
      if (GModeEclat = tmeRessTypA) then
        begin
        if (TobDet.GetValue('ACU_TYPEARTICLE')='PRE') then
          begin
            dSommeSolde0 := dSommeSolde0 + TobDet.GetValue('ACU_CUTOFFORIG');
            dSommeFAE0 := dSommeFAE0 + TobDet.GetValue('ACU_FAE');
            dSommeAAE0 := dSommeAAE0 + TobDet.GetValue('ACU_AAE');
            dSommePCA0 := dSommePCA0 + TobDet.GetValue('ACU_PCA');
          end
        else
        if (TobDet.GetValue('ACU_TYPEARTICLE')='FRA') then
          begin
            dSommeSolde1 := dSommeSolde1 + TobDet.GetValue('ACU_CUTOFFORIG');
            dSommeFAE1 := dSommeFAE1 + TobDet.GetValue('ACU_FAE');
            dSommeAAE1 := dSommeAAE1 + TobDet.GetValue('ACU_AAE');
            dSommePCA1 := dSommePCA1 + TobDet.GetValue('ACU_PCA');
          end
        else
        if (TobDet.GetValue('ACU_TYPEARTICLE')='MAR') then
          begin
            dSommeSolde2 := dSommeSolde2 + TobDet.GetValue('ACU_CUTOFFORIG');
            dSommeFAE2 := dSommeFAE2 + TobDet.GetValue('ACU_FAE');
            dSommeAAE2 := dSommeAAE2 + TobDet.GetValue('ACU_AAE');
            dSommePCA2 := dSommePCA2 + TobDet.GetValue('ACU_PCA');
          end;
        end;

      // Recherche des libellés des ressources
      if (GModeEclat = tmeRessArt) or (GModeEclat = tmeRessource) or (GModeEclat = tmeRessTypA) then
        begin
        iIndiceRess := RemplirTOBAssistant(TOBAssistant, AFOAssistants, TobDet.GetValue('ACU_RESSOURCE') );
        if (iIndiceRess >= 0) then
          TobDet.PutValue('LIBRESSOURCE', TOBAssistant.GetValue('ARS_LIBELLE'));
        end;

      // Recherche des libellés des articles
      if (GModeEclat = tmeRessArt) or (GModeEclat = tmeArticle) then
        begin
        TOBArt := FindTOBArtSaisAff(TOBArticles, TobDet.GetValue('ACU_ARTICLE'), true);
        if (TOBArt = Nil) then
          BEGIN
            TOBArt := CreerTOBArt(TOBArticles);
            TOBArt.InitValeurs (FALSE);
            RechArt := TrouverArticleSQL_GI (true, TobDet.GetValue('ACU_ARTICLE'), TOBArt, '');
            if (RechArt <> traOk) then
              begin
                TOBArt.Free;
                TOBArt := nil;
              end;
          END ;
        if (TOBArt <> Nil) then
          TobDet.PutValue('LIBARTICLE', TOBArt.GetValue('GA_LIBELLE'));
        end;

      // Recherche des libellés du type article
      if (GModeEclat = tmeRessTypA) or (GModeEclat = tmeTypeArticle) then
        TobDet.PutValue('LIBTYPEARTICLE', RechDom('GCTYPEARTICLE', TobDet.GetValue('ACU_TYPEARTICLE'), false));
      End;   // fin for

  THNumEdit(GetControl('TOTSOLDE')).Value:=dSommeSolde;
  THNumEdit(GetControl('TOTFAE')).Value:=dSommeFAE;
  THNumEdit(GetControl('TOTAAE')).Value:=dSommeAAE;
  THNumEdit(GetControl('TOTPCA')).Value:=dSommePCA;
  if (GModeEclat = tmeRessTypA) then
    begin
      THNumEdit(GetControl('TOTSOLDE0')).Value := dSommeSolde0;
      THNumEdit(GetControl('TOTFAE0')).Value := dSommeFAE0;
      THNumEdit(GetControl('TOTAAE0')).Value := dSommeAAE0;
      THNumEdit(GetControl('TOTPCA0')).Value := dSommePCA0;
      THNumEdit(GetControl('TOTSOLDE1')).Value := dSommeSolde1;
      THNumEdit(GetControl('TOTFAE1')).Value := dSommeFAE1;
      THNumEdit(GetControl('TOTAAE1')).Value := dSommeAAE1;
      THNumEdit(GetControl('TOTPCA1')).Value := dSommePCA1;
      THNumEdit(GetControl('TOTSOLDE2')).Value := dSommeSolde2;
      THNumEdit(GetControl('TOTFAE2')).Value := dSommeFAE2;
      THNumEdit(GetControl('TOTAAE2')).Value := dSommeAAE2;
      THNumEdit(GetControl('TOTPCA2')).Value := dSommePCA2;
    end;

  if (TForm(Ecran).Name = 'AFMODIFCUTOFFAMYO') then
    begin
      THNumEdit(GetControl('BASEVENTIL')).Value := dSommeSolde0 + dSommeSolde2;
      THNumEdit(GetControl('TOTFAR')).Value := dSommeFAR;
      THNumEdit(GetControl('TOTAAR')).Value := dSommeAAR;
      THNumEdit(GetControl('TOTCCA')).Value := dSommeCCA;
      THNumEdit(GetControl('TOTFAEACHAT')).Value := Abs(dSommeFAE0 + dSommeFAE2);
      THNumEdit(GetControl('TOTAAEACHAT')).Value := Abs(dSommeAAE0 + dSommeAAE2);
      THNumEdit(GetControl('TOTPCAACHAT')).Value := Abs(dSommePCA0 + dSommePCA2);
    end;

    // Remplissage du grid
  TobReg.PutGridDetail (GS, False, False, LesCol, true);


// PL le 20/02/03 : Spécif Amyot, l'écran est différent
if (TForm(Ecran).Name = 'AFMODIFCUTOFFAMYO') then
  begin
  // requête de la ligne d'achat
  ReqAchat := StringReplace (Req, '"CVE"', '"CAC"', [rfIgnoreCase]);

  // Création de la ligne d'achat
  TobAchat := Tob.Create( 'AFCUMUL', nil, -1);
  QQ := nil;
  try
    QQ := OpenSQL (ReqAchat, True);
    If Not QQ.EOF then
      begin
      TobAchat.SelectDB ('', QQ);
      bAchatOK := true;
      end;
  Finally
    Ferme(QQ);
  End;

  // S'il n'existe pas encore de ligne d'achat
  if (bAchatOK = false) and (TobReg.Detail[0] <> nil) then
    begin
    TobAchat.Dupliquer (TobReg.Detail[0], false, true);
    TobAchat.PutValue('ACU_TYPEAC', 'CAC');
    ClearLesMontantsSituation (TobAchat);
    TobAchat.PutValue('ACU_FAE', 0);
    TobAchat.PutValue('ACU_AAE', 0);
    TobAchat.PutValue('ACU_PCA', 0);
    TobAchat.PutValue('ACU_FAEDEV', 0);
    TobAchat.PutValue('ACU_AAEDEV', 0);
    TobAchat.PutValue('ACU_PCADEV', 0);
    // PL le 22/04/03 : le cut off d'achat est généré sur une fourniture : AF
    TobAchat.PutValue('ACU_TYPEARTICLE', 'MAR');
    TobAchat.PutValue('ACU_ARTICLE', '');
    TobAchat.PutValue('ACU_CODEARTICLE', '');
    // fin PL le 22/04/03
    end;

  // Affichage dans les zones d'achat
  THNumEdit(GetControl('TOTSOLDE2')).Value := dSommeSolde2; // PL le 24/04/03 : modif pour Amyot
  THNumEdit(GetControl('TOTFAR')).Value := TobAchat.GetValue('ACU_FAE');
  THNumEdit(GetControl('TOTAAR')).Value := TobAchat.GetValue('ACU_AAE');
  THNumEdit(GetControl('TOTCCA')).Value := TobAchat.GetValue('ACU_PCA');
  THNumEdit(GetControl('BASEVENTIL')).Value := dSommeSolde0 + dSommeSolde2 + THNumEdit(GetControl('TOTFAR')).Value - THNumEdit(GetControl('TOTAAR')).Value - THNumEdit(GetControl('TOTCCA')).Value;
  end;

end;


procedure TOF_AFMODIFCUTOFF.ReinitLesTotaux;
begin
  THNumEdit(GetControl('TOTSOLDE')).Value := 0;
  THNumEdit(GetControl('TOTFAE')).Value := 0;
  THNumEdit(GetControl('TOTAAE')).Value := 0;
  THNumEdit(GetControl('TOTPCA')).Value := 0;
  if (GModeEclat = tmeRessTypA) then
    begin
      THNumEdit(GetControl('TOTSOLDE0')).Value := 0;
      THNumEdit(GetControl('TOTFAE0')).Value := 0;
      THNumEdit(GetControl('TOTAAE0')).Value := 0;
      THNumEdit(GetControl('TOTPCA0')).Value := 0;
      THNumEdit(GetControl('TOTSOLDE1')).Value := 0;
      THNumEdit(GetControl('TOTFAE1')).Value := 0;
      THNumEdit(GetControl('TOTAAE1')).Value := 0;
      THNumEdit(GetControl('TOTPCA1')).Value := 0;
      THNumEdit(GetControl('TOTSOLDE2')).Value := 0;
      THNumEdit(GetControl('TOTFAE2')).Value := 0;
      THNumEdit(GetControl('TOTAAE2')).Value := 0;
      THNumEdit(GetControl('TOTPCA2')).Value := 0;
    end;

  if (TForm(Ecran).Name = 'AFMODIFCUTOFFAMYO') then
    begin
      THNumEdit(GetControl('BASEVENTIL')).Value := 0;
      THNumEdit(GetControl('TOTFAR')).Value := 0;
      THNumEdit(GetControl('TOTAAR')).Value := 0;
      THNumEdit(GetControl('TOTCCA')).Value := 0;
      THNumEdit(GetControl('TOTFAEACHAT')).Value := 0;
      THNumEdit(GetControl('TOTAAEACHAT')).Value := 0;
      THNumEdit(GetControl('TOTPCAACHAT')).Value := 0;
    end;
end;


procedure TOF_AFMODIFCUTOFF.Eclatement(AcbParle:boolean);
var
  dSommeSolde, dSommeFAE, dSommeAAE, dSommePCA : double;
  dSommeSolde0, dSommeFAE0, dSommeAAE0, dSommePCA0 : double;
  dSommeSolde1, dSommeFAE1, dSommeAAE1, dSommePCA1 : double;
  dSommeSolde2, dSommeFAE2, dSommeAAE2, dSommePCA2 : double;
  dSommeFAEAchat, dSommeAAEAchat, dSommePCAAchat : double;
  dRapportPresFour, dRapportFourPres : double;
begin
nextprevcontrol(Ecran);

dSommeSolde0 := 0;  dSommeSolde1 := 0;  dSommeSolde2 := 0;
dSommeFAE0 := 0;  dSommeFAE1 := 0;  dSommeFAE2 := 0;
dSommeAAE0 := 0;  dSommeAAE1 := 0;  dSommeAAE2 := 0;
dSommePCA0 := 0;  dSommePCA1 := 0;  dSommePCA2 := 0;
//dSommeFAEAchat := 0;  dSommeAAEAchat := 0;  dSommePCAAchat := 0; dRapportFourPres := 0; dRapportPresFour := 0;

dSommeSolde := THNumEdit(GetControl('TOTSOLDE')).Value;
if (GModeEclat = tmeRessTypA) then
  begin
  dSommeSolde0 := THNumEdit(GetControl('TOTSOLDE0')).Value;
  dSommeSolde1 := THNumEdit(GetControl('TOTSOLDE1')).Value;
  dSommeSolde2 := THNumEdit(GetControl('TOTSOLDE2')).Value;
  if  (dSommeSolde0 = 0) and (dSommeSolde1 = 0) and (dSommeSolde2 = 0) then exit;
  end
else
  if (dSommeSolde = 0) then exit;

// Message pour prevenir de l'ecrasement des donnees dans le grid
if AcbParle then
  if PGIAskAF('Cette opération va mettre à jour la partie Décision des cutoff.'+chr(13)+'Voulez-vous continuer ?',Ecran.Caption)<>mrYes then
    exit;

dSommeFAE := Arrondi (THNumEdit(GetControl('TOTFAE')).Value, V_PGI.OkDecV);
dSommeAAE := Arrondi (THNumEdit(GetControl('TOTAAE')).Value, V_PGI.OkDecV);
dSommePCA := Arrondi (THNumEdit(GetControl('TOTPCA')).Value, V_PGI.OkDecV);

bModified:=true;

if (GModeEclat = tmeRessTypA) then
  begin
  dSommeFAE1 := Arrondi (THNumEdit(GetControl('TOTFAE1')).Value, V_PGI.OkDecV);
  dSommeAAE1 := Arrondi (THNumEdit(GetControl('TOTAAE1')).Value, V_PGI.OkDecV);
  dSommePCA1 := Arrondi (THNumEdit(GetControl('TOTPCA1')).Value, V_PGI.OkDecV);

  // PL le 20/02/03 : Spécif Amyot, on réparti le total prestation + sous traitance au prorata de la production
  // sur les prestations et sur les fournitures
  if (TForm(Ecran).Name = 'AFMODIFCUTOFFAMYO') then
    begin
    dRapportPresFour := 1;
    dRapportFourPres := 1;
    if ((dSommeSolde0 + dSommeSolde2) <> 0) then
      begin
      dRapportPresFour := dSommeSolde0 / (dSommeSolde0 + dSommeSolde2);
      dRapportFourPres := dSommeSolde2 / (dSommeSolde0 + dSommeSolde2);
      end;

    dSommeFAEAchat := Arrondi (THNumEdit(GetControl('TOTFAEACHAT')).Value, V_PGI.OkDecV);
    dSommeAAEAchat := Arrondi (THNumEdit(GetControl('TOTAAEACHAT')).Value, V_PGI.OkDecV);
    dSommePCAAchat := Arrondi (THNumEdit(GetControl('TOTPCAACHAT')).Value, V_PGI.OkDecV);

    dSommeFAE0 := Arrondi (dSommeFAEAchat * dRapportPresFour, V_PGI.OkDecV);
    dSommeAAE0 := Arrondi (dSommeAAEAchat * dRapportPresFour, V_PGI.OkDecV);
    dSommePCA0 := Arrondi (dSommePCAAchat * dRapportPresFour, V_PGI.OkDecV);
    dSommeFAE2 := Arrondi (dSommeFAEAchat * dRapportFourPres, V_PGI.OkDecV);
    dSommeAAE2 := Arrondi (dSommeAAEAchat * dRapportFourPres, V_PGI.OkDecV);
    dSommePCA2 := Arrondi (dSommePCAAchat * dRapportFourPres, V_PGI.OkDecV);
    end
  else
    begin
    dSommeFAE0 := Arrondi (THNumEdit(GetControl('TOTFAE0')).Value, V_PGI.OkDecV);
    dSommeAAE0 := Arrondi (THNumEdit(GetControl('TOTAAE0')).Value, V_PGI.OkDecV);
    dSommePCA0 := Arrondi (THNumEdit(GetControl('TOTPCA0')).Value, V_PGI.OkDecV);
    dSommeFAE2 := Arrondi (THNumEdit(GetControl('TOTFAE2')).Value, V_PGI.OkDecV);
    dSommeAAE2 := Arrondi (THNumEdit(GetControl('TOTAAE2')).Value, V_PGI.OkDecV);
    dSommePCA2 := Arrondi (THNumEdit(GetControl('TOTPCA2')).Value, V_PGI.OkDecV);
    end;


  EclatementCutoff (dSommeSolde0, dSommeFAE0, dSommeAAE0, dSommePCA0, TobReg, GetControlText('ACU_AFFAIRE'), 'PRE');
  EclatementCutoff (dSommeSolde1, dSommeFAE1, dSommeAAE1, dSommePCA1, TobReg, GetControlText('ACU_AFFAIRE'), 'FRA');
  EclatementCutoff (dSommeSolde2, dSommeFAE2, dSommeAAE2, dSommePCA2, TobReg, GetControlText('ACU_AFFAIRE'), 'MAR');
  end
else
  begin
  EclatementCutoff (dSommeSolde, dSommeFAE, dSommeAAE, dSommePCA, TobReg, GetControlText('ACU_AFFAIRE'), '');
  end;

if (GModeEclat = tmeRessTypA) then
  begin
  THNumEdit(GetControl('TOTFAE0')).Value := dSommeFAE0;
  THNumEdit(GetControl('TOTAAE0')).Value := dSommeAAE0;
  THNumEdit(GetControl('TOTPCA0')).Value := dSommePCA0;
  THNumEdit(GetControl('TOTFAE1')).Value := dSommeFAE1;
  THNumEdit(GetControl('TOTAAE1')).Value := dSommeAAE1;
  THNumEdit(GetControl('TOTPCA1')).Value := dSommePCA1;
  THNumEdit(GetControl('TOTFAE2')).Value := dSommeFAE2;
  THNumEdit(GetControl('TOTAAE2')).Value := dSommeAAE2;
  THNumEdit(GetControl('TOTPCA2')).Value := dSommePCA2;
  THNumEdit(GetControl('TOTFAE')).Value := dSommeFAE0 + dSommeFAE1 + dSommeFAE2;
  THNumEdit(GetControl('TOTAAE')).Value := dSommeAAE0 + dSommeAAE1 + dSommeAAE2;
  THNumEdit(GetControl('TOTPCA')).Value := dSommePCA0 + dSommePCA1 + dSommePCA2;
  if (TForm(Ecran).Name = 'AFMODIFCUTOFFAMYO') then
    begin
    THNumEdit(GetControl('TOTFAEACHAT')).Value := dSommeFAE0 + dSommeFAE2;
    THNumEdit(GetControl('TOTAAEACHAT')).Value := dSommeAAE0 + dSommeAAE2;
    THNumEdit(GetControl('TOTPCAACHAT')).Value := dSommePCA0 + dSommePCA2;
    end;
  end
else
  begin
  THNumEdit(GetControl('TOTFAE')).Value := dSommeFAE;
  THNumEdit(GetControl('TOTAAE')).Value := dSommeAAE;
  THNumEdit(GetControl('TOTPCA')).Value := dSommePCA;
  end;

bTotModified:=false;

// Mise à jour du grid de saisie
TobReg.PutGridDetail (GS, False, False, LesCol, true);
end;

procedure TOF_AFMODIFCUTOFF.Insertion;
var
Arg, rep : string;
Critere, Champ, valeur  : String;
x : integer;
Crit1, Crit2, sFAE, sAAE, sPCA, Lib1, Lib2 :string;
TT : TOB;
begin
if (GModeEclat in [tmeGlobal, tmeSans]) then exit;
Crit1 := ''; Crit2 := ''; sFAE := ''; sAAE := ''; sPCA := ''; Lib1 := ''; Lib2 := '';
Arg := 'TYPE:CVE'; TT := nil;

rep:=AFLanceFiche_ModifCutOffAdd(Arg);
if (rep<>'') then
    begin
    // Recup des reponses
    Critere:=(Trim(ReadTokenSt(rep)));
    While (Critere <>'') do
        BEGIN
        if Critere<>'' then
            BEGIN
            x:=pos('=',Critere);
            if x<>0 then
                begin
                Champ:=copy(Critere,1,X-1);
                Valeur:=Copy (Critere,X+1,length(Critere)-X);
                end;

            if Champ = 'C1' then Crit1:=Valeur;
            if Champ = 'C2' then Crit2:=Valeur;
            if Champ = 'L1' then Lib1:=Valeur;
            if Champ = 'L2' then Lib2:=Valeur;
            if Champ = 'FAE' then sFAE:=Valeur;
            if Champ = 'AAE' then sAAE:=Valeur;
            if Champ = 'PCA' then sPCA:=Valeur;
            END;
        Critere:=(Trim(ReadTokenSt(rep)));
        END;

    // Test de l'existence de cet enregistrement
    case GModeEclat of
        tmeRessArt :begin
                TT := TobReg.FindFirst(['ACU_RESSOURCE','ACU_CODEARTICLE'], [Crit1, Crit2], true);
                end;
        tmeArticle :begin
                TT := TobReg.FindFirst(['ACU_CODEARTICLE'], [Crit2], true);
                end;
        tmeRessource :begin
                TT := TobReg.FindFirst(['ACU_RESSOURCE'], [Crit1], true);
                end;
        tmeRessTypA :begin
                TT := TobReg.FindFirst(['ACU_RESSOURCE','ACU_TYPEARTICLE'], [Crit1, Crit2], true);
                end;
        tmeTypeArticle :begin
                TT := TobReg.FindFirst(['ACU_TYPEARTICLE'], [Crit2], true);
                end;
    end;

    if (TT <> nil) then
        begin
        if (PGIAskAF('Cet enregistrement existe déjà, voulez-vous le mettre à jour avec les montants saisis ?', Ecran.Caption) <> mrYes) then
            exit;
        end
    else
        begin
        TT := TOB.Create ('AFCUMUML', TobReg, -1);
        TT.Dupliquer (TobReg.Detail[0], false, true);
        TT.AddChampSup ('LIBRESSOURCE', false);
        TT.AddChampSup ('LIBARTICLE', false);
        TT.AddChampSup ('LIBTYPEARTICLE', false);
        TT.AddChampSup ('LIGNEAJOUTEE', false);
        ClearLesMontantsSituation (TT);
        end;

    // Affectation des modifications  insertion
    case GModeEclat of
        tmeRessArt :begin
                TT.PutValue('ACU_RESSOURCE', Crit1);
                TT.PutValue('ACU_CODEARTICLE', Crit2);
                TT.PutValue('LIBRESSOURCE', Lib1);
                TT.PutValue('LIBARTICLE', Lib2);
                end;
        tmeArticle :begin
                TT.PutValue('ACU_CODEARTICLE', Crit2);
                TT.PutValue('LIBARTICLE', Lib2);
                end;
        tmeRessource :begin
                TT.PutValue('ACU_RESSOURCE', Crit1);
                TT.PutValue('LIBRESSOURCE', Lib1);
                end;
        tmeRessTypA :begin
                TT.PutValue('ACU_RESSOURCE', Crit1);
                TT.PutValue('LIBRESSOURCE', Lib1);
                TT.PutValue('ACU_TYPEARTICLE', Crit2);
                TT.PutValue('LIBTYPEARTICLE', Lib2);
                end;
        tmeTypeArticle :begin
                TT.PutValue('ACU_TYPEARTICLE', Crit2);
                TT.PutValue('LIBTYPEARTICLE', Lib2);
                end;
    end;
    TT.PutValue('ACU_FAE', sFAE);
    TT.PutValue('ACU_FAEDEV', sFAE);
    TT.PutValue('ACU_AAE', sAAE);
    TT.PutValue('ACU_AAEDEV', sAAE);
    TT.PutValue('ACU_PCA', sPCA);
    TT.PutValue('ACU_PCADEV', sPCA);

    // Mise à jour du grid de saisie
    TobReg.PutGridDetail (GS, False, False, LesCol, true);

    bModified := true;
//    nextprevcontrol (ecran);
    if (TobReg.detail.count <> 0) then
      begin
      PostMessage (GS.Handle, WM_KEYDOWN, VK_TAB, 0);
      application.processMessages;
      end;
    end;
end;


procedure TOF_AFMODIFCUTOFF.ClearLesMontantsSituation (TT:TOB);
begin
TT.PutValue('ACU_PVPROD', 0);
TT.PutValue('ACU_PVPRODDEV', 0);
TT.PutValue('ACU_PRPROD', 0);
TT.PutValue('ACU_PRPRODDEV', 0);
TT.PutValue('ACU_PVFACT', 0);
TT.PutValue('ACU_PVFACTDEV', 0);
TT.PutValue('ACU_QTEFACT', 0);
TT.PutValue('ACU_CUTOFFORIG', 0);
TT.PutValue('ACU_CUTOFFORIGDEV', 0);
end;

procedure TOF_AFMODIFCUTOFF.DessineTotaux;
var
Rect:TRect;
begin
TFVierge(Ecran).Hmtrad.ResizeGridColumns(GS) ;

if (TFVierge(Ecran).Name = 'AFMODIFCUTOFF') then
  begin
  Rect := GS.cellRect( ColSolde, 0);
  THNumEdit(GetControl('TOTSOLDE')).Left := Rect.Left;
  THNumEdit(GetControl('TOTSOLDE')).width := Rect.Right - Rect.Left;
  THNumEdit(GetControl('TOTSOLDE0')).Left := Rect.Left;
  THNumEdit(GetControl('TOTSOLDE0')).width := Rect.Right - Rect.Left;
  THNumEdit(GetControl('TOTSOLDE1')).Left := Rect.Left;
  THNumEdit(GetControl('TOTSOLDE1')).width := Rect.Right - Rect.Left;
  THNumEdit(GetControl('TOTSOLDE2')).Left := Rect.Left;
  THNumEdit(GetControl('TOTSOLDE2')).width := Rect.Right - Rect.Left;

  Rect := GS.cellRect( ColFAE, 0);
  THNumEdit(GetControl('TOTFAE')).Left := Rect.Left;
  THNumEdit(GetControl('TOTFAE')).width := Rect.Right - Rect.Left;
  THNumEdit(GetControl('TOTFAE0')).Left := Rect.Left;
  THNumEdit(GetControl('TOTFAE0')).width := Rect.Right - Rect.Left;
  THNumEdit(GetControl('TOTFAE1')).Left := Rect.Left;
  THNumEdit(GetControl('TOTFAE1')).width := Rect.Right - Rect.Left;
  THNumEdit(GetControl('TOTFAE2')).Left := Rect.Left;
  THNumEdit(GetControl('TOTFAE2')).width := Rect.Right - Rect.Left;

  Rect := GS.cellRect( ColAAE, 0);
  THNumEdit(GetControl('TOTAAE')).Left := Rect.Left;
  THNumEdit(GetControl('TOTAAE')).width := Rect.Right - Rect.Left;
  THNumEdit(GetControl('TOTAAE0')).Left := Rect.Left;
  THNumEdit(GetControl('TOTAAE0')).width := Rect.Right - Rect.Left;
  THNumEdit(GetControl('TOTAAE1')).Left := Rect.Left;
  THNumEdit(GetControl('TOTAAE1')).width := Rect.Right - Rect.Left;
  THNumEdit(GetControl('TOTAAE2')).Left := Rect.Left;
  THNumEdit(GetControl('TOTAAE2')).width := Rect.Right - Rect.Left;

  Rect := GS.cellRect( ColPCA, 0);
  THNumEdit(GetControl('TOTPCA')).Left := Rect.Left;
  THNumEdit(GetControl('TOTPCA')).width := Rect.Right - Rect.Left;
  THNumEdit(GetControl('TOTPCA0')).Left := Rect.Left;
  THNumEdit(GetControl('TOTPCA0')).width := Rect.Right - Rect.Left;
  THNumEdit(GetControl('TOTPCA1')).Left := Rect.Left;
  THNumEdit(GetControl('TOTPCA1')).width := Rect.Right - Rect.Left;
  THNumEdit(GetControl('TOTPCA2')).Left := Rect.Left;
  THNumEdit(GetControl('TOTPCA2')).width := Rect.Right - Rect.Left;

  THPanel(GetControl('PANELLIGNE')).Left := THLabel(GetControl('LBLSAISIEECLAT')).Left;
  THPanel(GetControl('PANELLIGNE')).width := THNumEdit(GetControl('TOTPCA0')).Left + THNumEdit(GetControl('TOTPCA0')).width - 4;
  end
else
if (TFVierge(Ecran).Name = 'AFMODIFCUTOFFAMYO') then
  begin
  Rect := GS.cellRect( ColSolde, 0);
  THNumEdit(GetControl('TOTSOLDE2')).Left := Rect.Left;
  THNumEdit(GetControl('TOTSOLDE2')).width := Rect.Right - Rect.Left;
  THNumEdit(GetControl('BASEVENTIL')).Left := Rect.Left;
  THNumEdit(GetControl('BASEVENTIL')).width := Rect.Right - Rect.Left;
  THNumEdit(GetControl('TOTSOLDE1')).Left := Rect.Left;
  THNumEdit(GetControl('TOTSOLDE1')).width := Rect.Right - Rect.Left;

  Rect := GS.cellRect( ColFAE, 0);
  THNumEdit(GetControl('TOTFAR')).Left := Rect.Left;
  THNumEdit(GetControl('TOTFAR')).width := Rect.Right - Rect.Left;
  THNumEdit(GetControl('LBLFAR')).Left := Rect.Left + trunc(THNumEdit(GetControl('TOTFAR')).width/2) - 10;
  THNumEdit(GetControl('TOTFAEACHAT')).Left := Rect.Left;
  THNumEdit(GetControl('TOTFAEACHAT')).width := Rect.Right - Rect.Left;
  THNumEdit(GetControl('LBLFAE')).Left := Rect.Left + trunc(THNumEdit(GetControl('TOTFAEACHAT')).width/2) - 10;
  THNumEdit(GetControl('TOTFAE1')).Left := Rect.Left;
  THNumEdit(GetControl('TOTFAE1')).width := Rect.Right - Rect.Left;

  Rect := GS.cellRect( ColAAE, 0);
  THNumEdit(GetControl('TOTAAR')).Left := Rect.Left;
  THNumEdit(GetControl('TOTAAR')).width := Rect.Right - Rect.Left;
  THNumEdit(GetControl('LBLAAR')).Left := Rect.Left + trunc(THNumEdit(GetControl('TOTAAR')).width/2) - 10;
  THNumEdit(GetControl('TOTAAEACHAT')).Left := Rect.Left;
  THNumEdit(GetControl('TOTAAEACHAT')).width := Rect.Right - Rect.Left;
  THNumEdit(GetControl('LBLAAE')).Left := Rect.Left + trunc(THNumEdit(GetControl('TOTAAEACHAT')).width/2) - 10;
  THNumEdit(GetControl('TOTAAE1')).Left := Rect.Left;
  THNumEdit(GetControl('TOTAAE1')).width := Rect.Right - Rect.Left;

  Rect := GS.cellRect( ColPCA, 0);
  THNumEdit(GetControl('TOTCCA')).Left := Rect.Left;
  THNumEdit(GetControl('TOTCCA')).width := Rect.Right - Rect.Left;
  THNumEdit(GetControl('LBLCCA')).Left := Rect.Left + trunc(THNumEdit(GetControl('TOTCCA')).width/2) - 10;
  THNumEdit(GetControl('TOTPCAACHAT')).Left := Rect.Left;
  THNumEdit(GetControl('TOTPCAACHAT')).width := Rect.Right - Rect.Left;
  THNumEdit(GetControl('LBLPCA')).Left := Rect.Left + trunc(THNumEdit(GetControl('TOTPCAACHAT')).width/2) - 10;
  THNumEdit(GetControl('TOTPCA1')).Left := Rect.Left;
  THNumEdit(GetControl('TOTPCA1')).width := Rect.Right - Rect.Left;

  THPanel(GetControl('PANELLIGNE')).Left := THLabel(GetControl('LBLCUTOFFACHAT')).Left;
  THPanel(GetControl('PANELLIGNE')).width := THNumEdit(GetControl('TOTCCA')).Left + THNumEdit(GetControl('TOTCCA')).width - 4;
  THPanel(GetControl('PANELLIGNE1')).Left := THLabel(GetControl('LBLCUTOFFACHAT')).Left;
  THPanel(GetControl('PANELLIGNE1')).width := THNumEdit(GetControl('TOTCCA')).Left + THNumEdit(GetControl('TOTCCA')).width - 4;
  end;

Rect := GS.cellRect( ColVide, 0);
THPanel(GetControl('PANEL3')).Left := Rect.Left;
THPanel(GetControl('PANEL3')).width := Rect.Right - Rect.Left;
// PL le 25/10/02 : ajout pour séparer la partie decision de la partie situation
Rect := GS.cellRect( ColSolde, 0);
THPanel(GetControl('PANEL3')).Left := Rect.Left;
THPanel(GetControl('PANEL3')).width := THPanel(GetControl('PANEL3')).width + Rect.Right - Rect.Left;

//Rect := GS.cellRect( ColVide, 0);
//THLabel(GetControl('LBLSITUATION')).Left := trunc(Rect.Left/2) - 40;
// PL le 07/01/03
//Rect := GS.cellRect( ColProd, 0);
//THLabel(GetControl('LBLSITUATION')).Left := trunc(Rect.Right/2) - 40;
Rect := GS.cellRect( ColSolde, 0);
THLabel(GetControl('LBLSITUATION')).Left := trunc(Rect.left/2) - 40;
////////////////// Fin PL le 25/10/02

Rect := GS.cellRect( ColVide, 0);
THLabel(GetControl('LBLDECISION')).Left := Rect.Right + trunc((THPanel(GetControl('PANEL2')).width - Rect.Right)/2)  - 35;
end;

Procedure TOF_AFMODIFCUTOFF.NomsChampsAffaire ( Var Aff,Aff0,Aff1,Aff2,Aff3,Aff4,Aff_,Aff0_,Aff1_,Aff2_,Aff3_,Aff4_,Tiers,Tiers_ : THEdit ) ;
BEGIN
Aff:=THEdit(GetControl('ACU_AFFAIRE'))   ; Aff0:=Nil ;
Aff1:=THEdit(GetControl('ACU_AFFAIRE1')) ;
Aff2:=THEdit(GetControl('ACU_AFFAIRE2')) ;
Aff3:=THEdit(GetControl('ACU_AFFAIRE3')) ;
Aff4:=THEdit(GetControl('ACU_AVENANT'))  ;
Tiers:=THEdit(GetControl('ACU_TIERS'))   ;
END ;


procedure TOF_AFMODIFCUTOFF.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_AFMODIFCUTOFF.OnDelete ;
  var
  TobDet : TOB;
begin
  //Inherited ;

  TobReg.GetGridDetail (GS, GS.RowCount - 1, 'AFCUMUL', LesCol);

  if (TobReg.detail.count <> 0) and (GS.Row > 1) then
    begin
    TobDet := TobReg.detail[ GS.Row - 1 ];
    if TobDet.FieldExists ('LIGNEAJOUTEE') then
      begin
        TobDet.Free;
        GS.DeleteRow (GS.Row);
        if (GS.RowCount > 1) then
          GS.Row := GS.Row - 1;
      end
    else
      begin
        PGIInfoAF ('Cette ligne ne peut être supprimée car elle a été créée par la génération du Cut off ou validée.', Ecran.Caption);
      end;
    end;
end ;

procedure TOF_AFMODIFCUTOFF.OnUpdate ;
begin
if (Ecran.activecontrol = GS) then
  begin
  if (TobReg.detail.count <> 0) then
    begin
    PostMessage (GS.Handle, WM_KEYDOWN, VK_TAB, 0);
    application.processMessages;
    end;
  end
else
  nextprevcontrol (Ecran);

LastError := 0;
EnErreur := False;
if TobReg.detail.count > 0 then
   begin
   if bTotModified then
        if PGIAskAF ('Les totaux ont été modifiés manuellement sans éclatement sur les lignes.' + chr(13) + 'Voulez-vous éclater ces totaux ?',Ecran.Caption)=mrYes then
            begin
            Eclatement (false);
            end;

   TobReg.GetGridDetail ( GS, GS.rowcount-1, 'AFCUMUL', LesCol);
   if bModified or bTotAchatModified then
      begin
      if PGIAskAF ('Voulez-vous enregistrer les modifications ?', Ecran.Caption) = mrYes then
         begin
           LastError := 0;
           //
           // Enregistrement des modifs dans la base
           TobReg.InsertOrUpdateDB (false);

           if (TobAchat <> nil ) then
              begin
              TobAchat.PutValue('ACU_FAE', THNumEdit(GetControl('TOTFAR')).Value);
              TobAchat.PutValue('ACU_FAEDEV', THNumEdit(GetControl('TOTFAR')).Value);
              TobAchat.PutValue('ACU_AAE', THNumEdit(GetControl('TOTAAR')).Value);
              TobAchat.PutValue('ACU_AAEDEV', THNumEdit(GetControl('TOTAAR')).Value);
              TobAchat.PutValue('ACU_PCA', THNumEdit(GetControl('TOTCCA')).Value);
              TobAchat.PutValue('ACU_PCADEV', THNumEdit(GetControl('TOTCCA')).Value);
              TobAchat.InsertOrUpdateDB (false);
              end;
           //
           if LastError <> 0 then exit;
           bModified := false;
           bTotAchatModified := false;
         end ;
      end;
   end;
LastError := 0;
end;

procedure TOF_AFMODIFCUTOFF.OnClose ;
begin

if bModified or bTotModified or bTotAchatModified then
  if PGIAskAF ('Etes-vous sûr de vouloir sortir sans valider les modifications ?', Ecran.Caption) <> mrYes then
    begin
      EnErreur := true;
      LastError := 1;
    end;

If EnErreur then begin LastError:=(-1); LastErrorMsg:=''; exit; end;

// Desallocations
TOBArt.Free; TOBArt := nil;
TOBArticles.Free; TOBArticles := nil;
TOBAssistant.Free ; TOBAssistant := Nil ;
AFOAssistants.Free; AFOAssistants := nil;
TobReg.free; TobReg := Nil;
TobAchat.free; TobAchat := nil;

  Inherited ;
end ;

procedure TOF_AFMODIFCUTOFF.OnLoad ;
begin
  Inherited ;
  EnErreur := False;
  bModified := false;

end ;

procedure TOF_AFMODIFCUTOFF.bFermerOnClick(SEnder: TObject);
begin
EnErreur:=false;
  inherited;
end;

function AFLanceFiche_ModifCutOff(Argument:String):string;
begin
if (GetParamSoc('SO_AFCLIENT') = cInClientAmyot) then
  Result:=AGLLanceFiche ('AFF','AFMODIFCUTOFFAMYO','','',Argument)
else
  Result:=AGLLanceFiche ('AFF','AFMODIFCUTOFF','','',Argument);

end;


procedure AGLEclatement(parms:array of variant; nb: integer ) ;
var  F : TForm ;
     MaTOF  : TOF;
begin
F:=TForm(Longint(Parms[0])) ;
if (F is TFvierge) then MaTOF:=TFvierge(F).LaTOF else exit;
if (MaTOF is TOF_AFMODIFCUTOFF) then TOF_AFMODIFCUTOFF(MaTOF).Eclatement(true) else exit;
end;

procedure AGLInsertion(parms:array of variant; nb: integer ) ;
var  F : TForm ;
     MaTOF  : TOF;
begin
F:=TForm(Longint(Parms[0])) ;
if (F is TFvierge) then MaTOF:=TFvierge(F).LaTOF else exit;
if (MaTOF is TOF_AFMODIFCUTOFF) then TOF_AFMODIFCUTOFF(MaTOF).Insertion else exit;
end;


Initialization
  registerclasses ( [ TOF_AFMODIFCUTOFF ] ) ;
  RegisterAglProc( 'Eclatement', True ,0, AGLEclatement);
  RegisterAglProc( 'Insertion', True ,0, AGLInsertion);
end.
