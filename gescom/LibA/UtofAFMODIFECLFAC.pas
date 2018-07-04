{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 12/08/2002
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : AFMODIFECLFAC ()
Mots clefs ... : TOF;AFMODIFECLFAC
*****************************************************************}
Unit UtofAFMODIFECLFAC ;

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
       db,  dbTables,FE_Main,
{$ENDIF}
     forms,M3FP,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HPanel,
     HMsgBox,
     HSysMenu,
     UTOF, utob, vierge,
     SaisUtil, dicoaf, FactUtil, utilCutoff,
     uafo_ressource,
     UTofAfBaseCodeAffaire,
     UtofAFMODIFCUTOFFADD ;

Type
  TOF_AFMODIFECLFAC = Class (TOF_AFBASECODEAFFAIRE)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnClose                  ; override ;
    procedure bFermerOnClick(SEnder: TObject);
    procedure bActuOnClick(SEnder: TObject);
    private
        GS : THGrid;
//        gCoefDevise, gOldPVFact : double;
        gCoefDevise : double;
        {Statut,}LesCol : String;
        nbCol, ColAff,ColArt,ColRess,ColTypA,{ColProd,}ColFact,{ColSolde,}ColVide,ColFactDev{,ColFAE,ColAAE,ColPCA} : integer;
        TobReg : TOB;
        bModified, bTotModified:boolean;
        procedure DessineTotaux;
        procedure GSCellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
//        procedure GSCellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
        procedure FormResize(Sender: TObject);
        procedure AfficheGrid;
        procedure TOTChange(Sender: TObject);
    public
        Action   :  TActionFiche ;
        EnErreur :  boolean;
        GModeEclat: T_ModeEclat;
        procedure NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2, Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers, Tiers_:THEdit); override;
        procedure Insertion;
        procedure ClearLesMontantsSituation(T:TOB);

  end ;

function AFLanceFiche_ModifEclatFact(Argument:String):string;

Implementation


procedure TOF_AFMODIFECLFAC.OnArgument (S : String ) ;
Var
    Critere, Champ, valeur, St  : String;
    x : integer;
begin
gCoefDevise:=1;
Action:=taModif ;
bModified:=false;
bTotModified:=false;
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

        if Champ = 'NATURE' then SetControlText('ACU_NATUREPIECEG',Valeur);
        if Champ = 'TIERS' then SetControlText('ACU_TIERS',Valeur);
        if Champ = 'AFFAIRE' then SetControlText('ACU_AFFAIRE',Valeur);
        if Champ = 'DATE' then SetControlText('ACU_DATE',Valeur);
        if Champ = 'NUMERO' then SetControlText('ACU_NUMERO',Valeur);
        if Champ = 'NUMECLAT' then SetControlText('ACU_NUMECLAT',Valeur);
        END;
    Critere:=(Trim(ReadTokenSt(S)));
    END;

if Action=taConsult then
   BEGIN
   FicheReadOnly(Ecran) ;
   END;




// Init du code affaire dans la tof ancêtre
Inherited ;

// Mode d'eclatement
GModeEclat:=DetermineModeEclatFact;
// Gestion du Grid
GS := THGRID(GetControl('GS'));
case GModeEclat of
   tmeRessArt :begin
                LesCol := 'ACU_AFFAIRE;ACU_RESSOURCE;ACU_CODEARTICLE';
                nbCol := 6;
                end;
   tmeRessource :begin
                LesCol := 'ACU_AFFAIRE;ACU_RESSOURCE';
                nbCol := 5;
                end;
   tmeRessTypA :begin
                LesCol := 'ACU_AFFAIRE;ACU_RESSOURCE;ACU_TYPEARTICLE';
                nbCol := 6;
                end;
end;
LesCol := LesCol + ';VIDE;ACU_PVFACT;ACU_PVFACTDEV';
GS.ColCount:=NbCol;

St:=LesCol;
for x:=0 to GS.ColCount-1 do
   BEGIN
   if x>2 then  GS.ColWidths[x]:=100;
   Champ:=ReadTokenSt(St) ;
   if Champ='ACU_RESSOURCE' then ColRess := x
   else if Champ='ACU_AFFAIRE' then ColAff := x
   else if Champ='ACU_CODEARTICLE' then ColArt := x
   else if Champ='ACU_TYPEARTICLE' then ColTypA := x
   else if Champ='ACU_PVFACT' then ColFact := x
   else if Champ='VIDE'  then ColVide:= x
   else if Champ='ACU_PVFACTDEV'  then ColFactDev:= x
   ;
   END ;

// dessin des deux ou trois premières colonnes
GS.ColWidths[ColAff]:=0;
GS.ColLengths[ColAff]:=-1;
case GModeEclat of
   tmeRessArt :begin
                GS.Cells[ColRess,0]:=TraduitGA('Ressource');
                GS.Cells[ColArt,0]:=TraduitGA('Article');
                GS.ColWidths[ColRess]:=150;
                GS.ColWidths[ColArt]:=150;
                GS.ColLengths[ColRess]:=-1;
                GS.ColLengths[ColArt]:=-1;
                end;
   tmeRessource :begin
                GS.Cells[ColRess,0]:=TraduitGA('Ressource');
                GS.ColWidths[ColRess]:=300;
                GS.ColLengths[ColRess]:=-1;
                end;
   tmeRessTypA :begin
                GS.Cells[ColRess,0]:=TraduitGA('Ressource');
                GS.Cells[ColTypA,0]:=TraduitGA('Type');
                GS.ColWidths[ColRess]:=200;
                GS.ColWidths[ColTypA]:=100;
                GS.ColLengths[ColRess]:=-1;
                GS.ColLengths[ColTypA]:=-1;
                end;
end;

// largeurs des colonnes
GS.ColWidths[ColFact]:=100;
GS.ColWidths[ColFactDev]:=100;
GS.ColWidths[ColVide]:=10;
// Protection des colonnes
//GS.ColLengths[ColFact]:=-1;
GS.ColLengths[ColFactDev]:=-1;
GS.ColLengths[ColVide]:=-1;
// justif des colonnes
GS.ColAligns[ColFact]:=taRightJustify;
GS.ColAligns[ColFactDev]:=taRightJustify;

// libellés des colonnes
GS.Cells[ColFact,0]:= TraduitGA('Facturé');
GS.Cells[ColFactDev,0]:= TraduitGA('Facturé en devise');
GS.Cells[ColVide,0]:= '';


// dessin des zones totaux aux bons emplacements sous les colonnes du grid associees
DessineTotaux;

// Chargement de la TOB pour affichage du GRID des cut off
AfficheGrid;

GS.OnCellExit := GSCellExit;
//GS.OnCellEnter := GSCellEnter;
TForm(Ecran).OnResize:=FormResize;
ChangeMask(THNumEdit(GetControl('TOTPVFACT')), V_PGI.OkDecV, '') ;
ChangeMask(THNumEdit(GetControl('TOTPVFACTDEV')), V_PGI.OkDecV, '') ;
THNumEdit(GetControl('TOTPVFACT')).onchange := TOTChange;

// Colonne courante dans le grid
GS.Col:= ColFact;
// Pour recadrage à doite correct dans le grid (mauvais fonctionnement du resize)
GS.ColWidths[ColVide]:=GS.ColWidths[ColVide]-3;

TToolBarButton97(GetControl('BFERME')).onclick := bFermerOnClick;
TToolBarButton97(GetControl('BACTU')).onclick := bActuOnClick;
end ;

procedure TOF_AFMODIFECLFAC.TOTChange(Sender: TObject);
var
  {dTotDiff,} dTotOld, dTotPvFact : double;
begin
  bTotModified:=true;
  //dTotDiff := THNumEdit(GetControl('TOTDIFF')).value;
  dTotOld := THNumEdit(GetControl('TOTOLD')).value;
  dTotPvFact := THNumEdit(GetControl('TOTPVFACT')).value;
  THNumEdit(GetControl('TOTDIFF')).value := Arrondi(dTotOld - dTotPvFact, V_PGI.OkDecV);
  //dTotDiff := THNumEdit(GetControl('TOTDIFF')).value;
end;


(*procedure TOF_AFMODIFECLFAC.GSCellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
begin
gOldPVFact := Valeur(GS.Cells[ACol,ARow]);
end;*)

procedure TOF_AFMODIFECLFAC.GSCellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
var
St,StC : String ;
//TOBL:TOB;
dSommePVFACT:double;
dSommePVFACTDEV:double;
begin
if Action=taConsult then Exit ;
//
// Gestion du format de saisie des cellules
//
St:=GS.Cells[ACol,ARow];
if (ACol in [ColFact]) then StC:=StrF00(Valeur(St),V_PGI.OkDecV) else
   StC:=St ;

GS.Cells[ACol,ARow]:=StC;

//if (gOldPVFact = Valeur(StC)) then exit;

//
// Gestion des sommes des colonnes du grid
//
// Récupération des données saisies dans le grid
TobReg.GetGridDetail(GS, GS.RowCount-1, 'AFCUMUL', LesCol );
dSommePVFACT:=TobReg.Somme('ACU_PVFACT', ['ACU_AFFAIRE'], [GetControlText('ACU_AFFAIRE')], true);
dSommePVFACTDEV:=TobReg.Somme('ACU_PVFACTDEV', ['ACU_AFFAIRE'], [GetControlText('ACU_AFFAIRE')], true);

// Affectation des PVFACT
if (TobReg<>nil) then
    begin
    TobReg.Detail[ARow-1].PutValue('ACU_PVFACTDEV', Arrondi(TobReg.Detail[ARow-1].GetValue('ACU_PVFACT')*gCoefDevise, V_PGI.OkDecV));
    end;

if (THNumEdit(GetControl('TOTPVFACT')).Value<>dSommePVFACT) then
    begin
    bModified:=true;
    TobReg.Detail[ARow-1].Putvalue('ACU_DATEMODIF', V_PGI.DateEntree);
    end;

THNumEdit(GetControl('TOTPVFACT')).Value:=dSommePVFACT;
THNumEdit(GetControl('TOTPVFACTDEV')).Value:=dSommePVFACTDEV;

end;

procedure TOF_AFMODIFECLFAC.FormResize(Sender: TObject);
begin
DessineTotaux;
end;

procedure TOF_AFMODIFECLFAC.AfficheGrid;
var
   TobDet : TOB;
   QQ : Tquery;
   Req{, Nat, Champ} : String;
   {x,} wi : Integer;
   dSommePVFACT, dSommePVFACTDev : double;

begin
  dSommePVFACT := 0; dSommePVFACTDev := 0;
  TobReg := Tob.Create('Liste Cutt Off',nil,-1);
// SELECT * : un enreg et nb de champs restreint
  Req := 'SELECT * FROM AFCUMUL Where ACU_TIERS="'+GetControlText('ACU_TIERS')
          + '" AND ACU_AFFAIRE="'+GetControlText('ACU_AFFAIRE')
          + '" AND ACU_NATUREPIECEG="'+GetControlText('ACU_NATUREPIECEG')
          + '" AND ACU_NUMECLAT='+GetControlText('ACU_NUMECLAT')
          + ' AND ACU_NUMERO='+GetControlText('ACU_NUMERO')
          + ' AND ACU_DATE="'+usdatetime(strtodate(GetControlText('ACU_DATE')));

  case GModeEclat of
     tmeRessArt :begin
                  Req := Req + '" ORDER BY ACU_RESSOURCE,ACU_CODEARTICLE';
                  end;
     tmeRessource :begin
                  Req := Req + '" ORDER BY ACU_RESSOURCE';
                  end;
     tmeRessTypA :begin
                  Req := Req + '" ORDER BY ACU_RESSOURCE,ACU_TYPEARTICLE';
                  end;
      else
                  begin
                  Req := Req + '" ORDER BY ACU_AFFAIRE';
                  end;
  end;

  QQ := nil;
  try
    QQ := OpenSQL(Req,True) ;
    If Not QQ.EOF then TobReg.LoadDetailDB('AFCUMUL','','',QQ,True) Else Exit;
  Finally
    Ferme(QQ);
    GS.RowCount:=TobReg.Detail.count+1;
  End;

  gCoefDevise:=1;
  if (TobReg.Detail.Count > 0) Then
    for wi:=0 To TobReg.Detail.Count-1 do
      Begin
      TobDet := TobReg.Detail[wi];
      dSommePVFACT:=dSommePVFACT+TobDet.GetValue('ACU_PVFACT');
      dSommePVFACTDev:=dSommePVFACTDev+TobDet.GetValue('ACU_PVFACTDEV');
      if (gCoefDevise=1) and (TobDet.GetValue('ACU_PVFACT')<>0) and (TobDet.GetValue('ACU_PVFACT')<>TobDet.GetValue('ACU_PVFACTDEV')) then
          begin
          gCoefDevise:=TobDet.GetValue('ACU_PVFACTDEV')/TobDet.GetValue('ACU_PVFACT');
          end;
      End;   // fin for

  THNumEdit(GetControl('TOTOLD')).Value:=dSommePVFACT;
  THNumEdit(GetControl('TOTPVFACT')).Value:=dSommePVFACT;
  THNumEdit(GetControl('TOTPVFACTDEV')).Value:=dSommePVFACTDev;
  THNumEdit(GetControl('TOTDIFF')).Value:=0;

  TobReg.PutGridDetail(GS,False,False,LesCol,true);
end;

procedure TOF_AFMODIFECLFAC.Insertion;
var
  Arg, rep :string;
  Critere, Champ, valeur  : String;
  x {, Index} : integer;
  Crit1, Crit2, sPVFACT:string;
  T {, TOBArt}:TOB;
begin
  if (GModeEclat in [tmeGlobal, tmeSans]) then exit;
  Crit1:=''; Crit2:=''; sPVFACT:='';
  Arg := 'TYPE:FAC'; T := nil;

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
              if Champ = 'PVFACT' then sPVFACT:=Valeur;
              END;
          Critere:=(Trim(ReadTokenSt(rep)));
          END;

      // Test de l'existence de cet enregistrement
      case GModeEclat of
          tmeRessArt :begin
                  T:=TobReg.FindFirst(['ACU_RESSOURCE','ACU_CODEARTICLE'], [Crit1, Crit2], true);
                  end;
          tmeRessource :begin
                  T:=TobReg.FindFirst(['ACU_RESSOURCE'], [Crit1], true);
                  end;
          tmeRessTypA :begin
                  T:=TobReg.FindFirst(['ACU_RESSOURCE','ACU_TYPEARTICLE'], [Crit1, Crit2], true);
                  end;
      end;

      if (T <> nil) then
          begin
          if PGIAskAF('Cet enregistrement existe déjà, voulez-vous le mettre à jour avec le montant saisi ?',Ecran.Caption)<>mrYes then
              exit;
          end
      else
          begin
          T := TOB.Create('AFCUMUML',TobReg,-1) ;
          T.Dupliquer(TobReg.Detail[0], false, true);
          ClearLesMontantsSituation(T);
          end;

      // Affectation des modifications  insertion
      case GModeEclat of
          tmeRessArt :begin
                  T.PutValue('ACU_RESSOURCE', Crit1);
                  T.PutValue('ACU_CODEARTICLE', Crit2);
                  end;
          tmeRessource :begin
                  T.PutValue('ACU_RESSOURCE', Crit1);
                  end;
          tmeRessTypA :begin
                  T.PutValue('ACU_RESSOURCE', Crit1);
                  T.PutValue('ACU_TYPEARTICLE', Crit2);
                  end;
      end;
      T.PutValue('ACU_PVFACT', sPVFACT);
      T.PutValue('ACU_PVFACTDEV', Arrondi(T.GetValue('ACU_PVFACT')*gCoefDevise, V_PGI.OkDecV));

      // Mise à jour du grid de saisie
      TobReg.PutGridDetail(GS,False,False,LesCol,true);

      bModified:=true;

      THNumEdit(GetControl('TOTPVFACT')).Value:=TobReg.Somme('ACU_PVFACT', ['ACU_AFFAIRE'], [GetControlText('ACU_AFFAIRE')], true);
      THNumEdit(GetControl('TOTPVFACTDEV')).Value:=TobReg.Somme('ACU_PVFACTDEV', ['ACU_AFFAIRE'], [GetControlText('ACU_AFFAIRE')], true);
      end;
end;


procedure TOF_AFMODIFECLFAC.ClearLesMontantsSituation(T:TOB);
begin
T.PutValue('ACU_PVFACT', 0);
T.PutValue('ACU_PVFACTDEV', 0);
end;

procedure TOF_AFMODIFECLFAC.DessineTotaux;
var
Rect:TRect;
begin
TFVierge(Ecran).Hmtrad.ResizeGridColumns(GS) ;

Rect := GS.cellRect( ColFact, 0);
THNumEdit(GetControl('TOTPVFACT')).Left := Rect.Left;
THNumEdit(GetControl('TOTPVFACT')).width := Rect.Right - Rect.Left;

THNumEdit(GetControl('TOTDIFF')).Left := THNumEdit(GetControl('TOTPVFACT')).Left;
THNumEdit(GetControl('TOTDIFF')).width := THNumEdit(GetControl('TOTPVFACT')).width;
THNumEdit(GetControl('LBLDIFF')).Left := THNumEdit(GetControl('TOTPVFACT')).Left - THNumEdit(GetControl('LBLDIFF')).width - 5;

THNumEdit(GetControl('TOTOLD')).Left := Rect.Left - THNumEdit(GetControl('TOTPVFACT')).width - 5;
THNumEdit(GetControl('TOTOLD')).width := THNumEdit(GetControl('TOTPVFACT')).width;

Rect := GS.cellRect( ColFactDev, 0);
THNumEdit(GetControl('TOTPVFACTDEV')).Left := Rect.Left;
THNumEdit(GetControl('TOTPVFACTDEV')).width := Rect.Right - Rect.Left;

Rect := GS.cellRect( ColVide, 0);
THPanel(GetControl('PANEL3')).Left := Rect.Left;
THPanel(GetControl('PANEL3')).width := Rect.Right - Rect.Left;

Rect := GS.cellRect( ColVide, 0);
THLabel(GetControl('LBLSITUATION')).Left := trunc(Rect.Left/2) - 40;

Rect := GS.cellRect( ColVide, 0);
THLabel(GetControl('LBLDECISION')).Left := Rect.Right + trunc((THPanel(GetControl('PANEL2')).width - Rect.Right)/2)  - 35;
end;

Procedure TOF_AFMODIFECLFAC.NomsChampsAffaire ( Var Aff,Aff0,Aff1,Aff2,Aff3,Aff4,Aff_,Aff0_,Aff1_,Aff2_,Aff3_,Aff4_,Tiers,Tiers_ : THEdit ) ;
BEGIN
Aff:=THEdit(GetControl('ACU_AFFAIRE'))   ; Aff0:=Nil ;
Aff1:=THEdit(GetControl('ACU_AFFAIRE1')) ;
Aff2:=THEdit(GetControl('ACU_AFFAIRE2')) ;
Aff3:=THEdit(GetControl('ACU_AFFAIRE3')) ;
Aff4:=THEdit(GetControl('ACU_AVENANT'))  ;
Tiers:=THEdit(GetControl('ACU_TIERS'))   ;
END ;


procedure TOF_AFMODIFECLFAC.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_AFMODIFECLFAC.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_AFMODIFECLFAC.OnUpdate ;
begin
nextprevcontrol(ecran);
LastError:=0;
EnErreur:=False;
if TobReg.detail.count >0 then
   begin
   if (THNumEdit(GetControl('TOTDIFF')).value<>0) then
        begin
        PGIInfoAF('Validation impossible : le total de vos factures doit rester inchangé. ' + floattostr(THNumEdit(GetControl('TOTDIFF')).value),Ecran.Caption);
        EnErreur:=true;
        LastError:=1;
        exit;
        end;

   TobReg.GetGridDetail( GS,GS.rowcount-1,'AFCUMUL',LesCol) ;
   if bModified  then
      begin
      if PGIAskAF('Voulez-vous enregistrer les modifications ?',Ecran.Caption)=mrYes then
         begin
         LastError:=0;
         //
         // Enregistrement des modifs dans la base
         TobReg.InsertOrUpdateDB(false);
         //
         if LastError<>0 then exit;
         bModified:=false;
         end ;
      end;
   end;
LastError:=0;
end;

procedure TOF_AFMODIFECLFAC.OnClose ;
begin

If EnErreur then begin LastError:=(-1); LastErrorMsg:=''; exit; end;

// Desallocations
TobReg.free;TobReg:=Nil;

  Inherited ;
end ;

procedure TOF_AFMODIFECLFAC.OnLoad ;
begin
  Inherited ;
EnErreur:=False;
bModified:=false;

end ;

procedure TOF_AFMODIFECLFAC.bFermerOnClick(SEnder: TObject);
begin
EnErreur := false;
  inherited;
end;
// mcd 05/03/03 ajout du bouton qui permet d'affecter automatiquement le mtt
// de la zone différence dans la ligne en cours.
procedure TOF_AFMODIFECLFAC.bActuOnClick(SEnder: TObject);
Var mtt:double;
begin
  Mtt := TobReg.Detail[GS.Row-1].GetValue('ACU_PVFACT');
  Mtt := Mtt+ THNumEdit(GetControl('TOTDIFF')).value;
  TobReg.Detail[GS.Row-1].PutValue('ACU_PVFACT', Arrondi(Mtt, V_PGI.OkDecV));
  TobReg.Detail[GS.Row-1].PutValue('ACU_PVFACTDEV', Arrondi(TobReg.Detail[GS.Row-1].GetValue('ACU_PVFACT')*gCoefDevise, V_PGI.OkDecV));
    // Mise à jour du grid de saisie
  TobReg.PutGridDetail (GS, False, False, LesCol, true);
  bModified := true;
  PostMessage (GS.Handle, WM_KEYDOWN, VK_TAB, 0);
  application.processMessages;

end;

function AFLanceFiche_ModifEclatFact(Argument:String):string;
begin
Result:=AGLLanceFiche ('AFF','AFMODIFECLFAC','','',Argument);
end;


procedure AGLInsertionFact(parms:array of variant; nb: integer ) ;
var  F : TForm ;
     MaTOF  : TOF;
begin
F:=TForm(Longint(Parms[0])) ;
if (F is TFvierge) then MaTOF:=TFvierge(F).LaTOF else exit;
if (MaTOF is TOF_AFMODIFECLFAC) then TOF_AFMODIFECLFAC(MaTOF).Insertion else exit;
end;


Initialization
  registerclasses ( [ TOF_AFMODIFECLFAC ] ) ;
  RegisterAglProc( 'InsertionFac', True ,0, AGLInsertionFact);
end.
