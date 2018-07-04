{***********UNITE*************************************************
Auteur  ...... :  DESSEIGNET                           
Créé le ...... : 22/08/2003
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : AFMODIFBONI ()
Mots clefs ... : TOF;AFMODIFBONI
*****************************************************************}
Unit UtofAFMODIFBONI ;

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
     HStatus,
     UTOF, utob, vierge, dicoaf,
     SaisUtil,
     UtofAfModifBoni_Add, ActiviteUtil, UtilArticle,
     UTofAfBaseCodeAffaire;

Type
  TOF_AFMODIFBONI = Class (TOF_AFBASECODEAFFAIRE)
    procedure OnUpdate                 ; override ;
    procedure  OnLoad                  ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnClose                  ; override ;
    procedure bFermerOnClick(SEnder: TObject);
    procedure bActuOnClick(SEnder: TObject);
    private
        GS : THGrid;
        gCoefDevise : double;
        LesCol : String;
        nbCol, ColArt,ColRess,ColBoni,ColVide : integer;
        TobReg : TOB;
        bModified, bTotModified:boolean;
        procedure DessineTotaux;
        procedure GSCellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
        procedure FormResize(Sender: TObject);
        procedure AfficheGrid;
        procedure TOTChange(Sender: TObject);
        procedure ValideLesBoniMali;
    public
        Action   :  TActionFiche ;
        EnErreur :  boolean;
        procedure NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2, Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers, Tiers_:THEdit); override;
        procedure Insertion;
        procedure ClearLesMontantsSituation(T:TOB);

  end ;

function AFLanceFiche_ModifBoni(Argument:String):string;

Implementation


procedure TOF_AFMODIFBONI.OnArgument (S : String ) ;
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
        if Champ = 'TIERS' then SetControlText('ACT_TIERS',Valeur)
        else if Champ = 'AFFAIRE' then SetControlText('ACT_AFFAIRE',Valeur)
        else if Champ = 'DATE' then SetControlText('ACT_DATEACTIVITE',Valeur)
        else if Champ = 'NUM' then SetControlText('ACT_NUMAPPREC',Valeur);
        END;
    Critere:=(Trim(ReadTokenSt(S)));
    END;

if Action=taConsult then
   BEGIN
   FicheReadOnly(Ecran) ;
   END;

// Init du code affaire dans la tof ancêtre
Inherited ;

// Force la protection des champs Affaire
EditAff.ReadOnly := true;
EditAff1.ReadOnly := true;
EditAff2.ReadOnly := true;
EditAff3.ReadOnly := true;
EditAff4.ReadOnly := true;
SetControlProperty ('ACT_DATEACTIVITE', 'ReadOnly', true);

// Gestion du Grid
GS := THGRID(GetControl('GS'));
LesCol := 'ACT_RESSOURCE;ACT_CODEARTICLE';
nbCol := 4;
LesCol := LesCol + ';VIDE;ACT_TOTVENTE';
GS.ColCount:=NbCol;

St:=LesCol;
for x:=0 to GS.ColCount-1 do
   BEGIN
   if x>2 then  GS.ColWidths[x]:=100;
   Champ:=ReadTokenSt(St) ;
   if Champ='ACT_RESSOURCE' then ColRess := x
   else if Champ='ACT_CODEARTICLE' then ColArt := x
   else if Champ='ACT_TOTVENTE' then ColBoni := x
   else if Champ='VIDE'  then ColVide:= x
   ;
   END ;

// dessin des deux ou trois premières colonnes
GS.Cells[ColRess,0]:=TraduitGA('Ressource');
GS.Cells[ColArt,0]:=TraduitGA('Article');
GS.ColWidths[ColRess]:=150;
GS.ColWidths[ColArt]:=150;
GS.ColLengths[ColRess]:=-1;
GS.ColLengths[ColArt]:=-1;

// largeurs des colonnes
GS.ColWidths[ColBoni]:=150;
GS.ColWidths[ColVide]:=10;
// Protection des colonnes
GS.ColLengths[ColVide]:=-1;
// justif des colonnes
GS.ColAligns[ColBoni]:=taRightJustify;

// libellés des colonnes
GS.Cells[ColBoni,0]:= TraduitGA('Boni/Mali');
GS.Cells[ColVide,0]:= '';

// dessin des zones totaux aux bons emplacements sous les colonnes du grid associees
DessineTotaux;
// Chargement de la TOB pour affichage du GRID des cut off
AfficheGrid;

GS.OnCellExit := GSCellExit;
//GS.OnCellEnter := GSCellEnter;
TForm(Ecran).OnResize:=FormResize;
ChangeMask(THNumEdit(GetControl('TOTBONI')), V_PGI.OkDecV, '') ;
THNumEdit(GetControl('TOTBONI')).onchange := TOTChange;

// Colonne courante dans le grid
GS.Col:= ColBoni;
// Pour recadrage à doite correct dans le grid (mauvais fonctionnement du resize)
GS.ColWidths[ColVide]:=GS.ColWidths[ColVide]-3;

TToolBarButton97(GetControl('BFERME')).onclick := bFermerOnClick;
TToolBarButton97(GetControl('BACTU')).onclick := bActuOnClick;
end ;

procedure TOF_AFMODIFBONI.TOTChange(Sender: TObject);
var
  dTotOld, dTotBoni : double;
begin
  bTotModified:=true;
  dTotOld := THNumEdit(GetControl('TOTOLD')).value;
  dTotBoni := THNumEdit(GetControl('TOTBONI')).value;
  THNumEdit(GetControl('TOTDIFF')).value := Arrondi(dTotOld - dTotBoni, V_PGI.OkDecV);
end;



procedure TOF_AFMODIFBONI.GSCellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
var
St,StC : String ;
dSommeBoni:double;
begin
if Action=taConsult then Exit ;
//
// Gestion du format de saisie des cellules
//
St:=GS.Cells[ACol,ARow];
if (ACol in [ColBoni]) then StC:=StrF00(Valeur(St),V_PGI.OkDecV) else
   StC:=St ;

GS.Cells[ACol,ARow]:=StC;

//
// Gestion des sommes des colonnes du grid
//
// Récupération des données saisies dans le grid
TobReg.GetGridDetail(GS, GS.RowCount-1, 'ACTIVITE', LesCol );
dSommeBoni:=TobReg.Somme('ACT_TOTVENTE', ['ACT_AFFAIRE'], [GetControlText('ACT_AFFAIRE')], true);

if (THNumEdit(GetControl('TOTBONI')).Value<>dSommeBoni) then
    begin
    bModified:=true;
    TobReg.Detail[ARow-1].Putvalue('ACT_DATEMODIF', V_PGI.DateEntree);
    TobReg.detail[Arow-1].putvalue('ACT_PUVENTE', TobReg.detail[Arow-1].Getvalue('ACT_TOTVENTE'));
    end;

THNumEdit(GetControl('TOTBoni')).Value:=dSommeBoni;

end;

procedure TOF_AFMODIFBONI.FormResize(Sender: TObject);
begin
DessineTotaux;
end;

procedure TOF_AFMODIFBONI.AfficheGrid;
var
   TobDet : TOB;
   QQ : Tquery;
   Req : String;
    wi : Integer;
   dSommeBoni : double;

begin
  dSommeBoni := 0;
  TobReg := Tob.Create('Liste Boni',nil,-1);
// SELECT * : un enreg et nb de champs restreint
  Req := 'SELECT * FROM ACTIVITE Where '
          + ' ACT_AFFAIRE="'+GetControlText('ACT_AFFAIRE')
          + '" AND ACT_NUMAPPREC='+ GetControlText('ACT_NUMAPPREC')
          + ' AND ACT_DATEACTIVITE="'+usdatetime(strtodate(GetControlText('ACT_DATEACTIVITE'))) +'"';
  Req := Req + ' ORDER BY ACT_RESSOURCE,ACT_CODEARTICLE';
  QQ := nil;
  try
    QQ := OpenSQL(Req,True) ;
    If Not QQ.EOF then TobReg.LoadDetailDB('ACTIVITE','','',QQ,True) Else Exit;
  Finally
    Ferme(QQ);
    GS.RowCount:=TobReg.Detail.count+1;
  End;

  gCoefDevise:=1;
  if (TobReg.Detail.Count > 0) Then
    for wi:=0 To TobReg.Detail.Count-1 do
      Begin
      TobDet := TobReg.Detail[wi];
      dSommeBoni:=dSommeBoni+TobDet.GetValue('ACT_TOTVENTE');
      End;   // fin for

  THNumEdit(GetControl('TOTOLD')).Value:=dSommeBoni;
  THNumEdit(GetControl('TOTBoni')).Value:=dSommeBoni;
  THNumEdit(GetControl('TOTDIFF')).Value:=0;

  TobReg.PutGridDetail(GS,False,False,LesCol,true);
end;

procedure TOF_AFMODIFBONI.Insertion;
var
  rep :string;
  Critere, Champ, valeur  : String;
  x  : integer;
  Crit1, Crit2, sBoni:string;
  T :TOB;
begin
  Crit1:=''; Crit2:=''; sBoni:='';
  rep:=AFLanceFiche_ModifBoni_Add('');
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
              if Champ = 'BONI' then sBoni :=Valeur;
              END;
          Critere:=(Trim(ReadTokenSt(rep)));
          END;
      // Test de l'existence de cet enregistrement
      T:=TobReg.FindFirst(['ACT_RESSOURCE','ACT_CODEARTICLE'], [Crit1, Crit2], true);
      if (T <> nil) then
          begin
          if PGIAskAF('Cet enregistrement existe déjà, voulez-vous le mettre à jour avec le montant saisi ?',Ecran.Caption)<>mrYes then
              exit;
          end
      else
          begin
          T := TOB.Create('ACTIVITE',TobReg,-1) ;
          T.Dupliquer(TobReg.Detail[0], false, true);
          ClearLesMontantsSituation(T);
          end;
      // Affectation des modifications  insertion
      T.PutValue('ACT_RESSOURCE', Crit1);
      T.PutValue('ACT_CODEARTICLE', Crit2);
      T.PUtValue('ACT_ARTICLE', CodeArticleUnique(crit2,'','','','',''));
      T.PutValue('ACT_TOTVENTE', sBoni);
      T.PutValue('ACT_PUVENTE', sBoni);
      T.PutValue ('ACT_DATECREATION', Nowh);
      // Mise à jour du grid de saisie
      TobReg.PutGridDetail(GS,False,False,LesCol,true);
      bModified:=true;
      THNumEdit(GetControl('TOTBONI')).Value:=TobReg.Somme('ACT_TOTVENTE', ['ACT_AFFAIRE'], [GetControlText('ACT_AFFAIRE')], true);
      end;
end;


procedure TOF_AFMODIFBONI.ClearLesMontantsSituation(T:TOB);
begin
T.PutValue('ACT_TOTVENTE', 0);
T.PutValue('ACT_PUVENTE', 0);
end;

procedure TOF_AFMODIFBONI.DessineTotaux;
var
Rect:TRect;
begin
TFVierge(Ecran).Hmtrad.ResizeGridColumns(GS) ;

Rect := GS.cellRect( ColBoni, 0);
THNumEdit(GetControl('TOTBONI')).Left := Rect.Left;
THNumEdit(GetControl('TOTBONI')).width := Rect.Right - Rect.Left;

THNumEdit(GetControl('TOTDIFF')).Left := THNumEdit(GetControl('TOTBONI')).Left;
THNumEdit(GetControl('TOTDIFF')).width := THNumEdit(GetControl('TOTBONI')).width;
THNumEdit(GetControl('LBLDIFF')).Left := THNumEdit(GetControl('TOTBONI')).Left - THNumEdit(GetControl('LBLDIFF')).width - 5;

THNumEdit(GetControl('TOTOLD')).Left := Rect.Left - THNumEdit(GetControl('TOTBONI')).width - 5;
THNumEdit(GetControl('TOTOLD')).width := THNumEdit(GetControl('TOTBONI')).width;


Rect := GS.cellRect( ColVide, 0);
THPanel(GetControl('PANEL3')).Left := Rect.Left;
THPanel(GetControl('PANEL3')).width := Rect.Right - Rect.Left;

Rect := GS.cellRect( ColVide, 0);
THLabel(GetControl('LBLSITUATION')).Left := trunc(Rect.Left/2) - 40;

Rect := GS.cellRect( ColVide, 0);
THLabel(GetControl('LBLDECISION')).Left := Rect.Right + trunc((THPanel(GetControl('PANEL2')).width - Rect.Right)/2)  - 35;
end;

Procedure TOF_AFMODIFBONI.NomsChampsAffaire ( Var Aff,Aff0,Aff1,Aff2,Aff3,Aff4,Aff_,Aff0_,Aff1_,Aff2_,Aff3_,Aff4_,Tiers,Tiers_ : THEdit ) ;
BEGIN
Aff:=THEdit(GetControl('ACT_AFFAIRE'))   ; Aff0:=Nil ;
Aff1:=THEdit(GetControl('ACT_AFFAIRE1')) ;
Aff2:=THEdit(GetControl('ACT_AFFAIRE2')) ;
Aff3:=THEdit(GetControl('ACT_AFFAIRE3')) ;
Aff4:=THEdit(GetControl('ACT_AVENANT'))  ;
Tiers:=THEdit(GetControl('ACT_TIERS'))   ;
END ;




procedure TOF_AFMODIFBONI.OnUpdate ;
var
    ResTrans : TIOErr;
begin
nextprevcontrol(ecran);
LastError:=0;
EnErreur:=False;
if TobReg.detail.count >0 then
   begin
   if (THNumEdit(GetControl('TOTDIFF')).value<>0) then
        begin
        PGIInfoAF('Validation impossible : le total de vos Boni/mali doit rester inchangé. ' + floattostr(THNumEdit(GetControl('TOTDIFF')).value),Ecran.Caption);
        EnErreur:=true;
        LastError:=1;
        exit;
        end;
   TobReg.GetGridDetail( GS,GS.rowcount-1,'AFACTIVITE',LesCol) ;
   if bModified  then
      begin
      if PGIAskAF ('Voulez-vous enregistrer les modifications ?', Ecran.Caption) = mrYes then
         begin
         LastError:=0;
         //

          ResTrans := Transactions (ValideLesBoniMali, 2);

          Case ResTrans of
            oeOk :      begin
                          bModified := false;
                        end;
            oeUnknown : begin
                          PGIBoxAf ('ATTENTION : modifications non enregistrées.', Ecran.Caption);
                          LastError := 1;
                          Exit;
                        end ;
          end ;
         end ;
      end;
   end;
LastError:=0;
end;

procedure TOF_AFMODIFBONI.ValideLesBoniMali;
var
  Sql : string;
  TobDet : Tob;
  Num, ii : integer;
begin
  try
    SourisSablier;
    InitMove (2, '');

    // suppression ancien enrgt de la base
    TobDet := TobReg.detail[1];
    Sql := 'DELETE ACTIVITE WHERE ACT_TYPEACTIVITE="BON" and';
    Sql := Sql + ' ACT_AFFAIRE ="' + TobDet.Getvalue('ACT_AFFAIRE');
    Sql := Sql + ' " AND ACT_NUMAPPREC =' + IntToStr (TobDet.Getvalue('ACT_NUMAPPREC'));
    ExecuteSql (Sql);

    MoveCur (False);

    if V_PGI.IoError = oeOk then
      begin
        //recherche n° unique
        Num := ProchainPlusNumLigneUniqueActivite ('BON', TobDet.Getvalue('ACT_AFFAIRE'));
        For ii := 0 to TobReg.detail.count - 1 do
          begin
            TobDet := TobReg.detail[ii];
            TobDet.PutValue ('ACT_NUMLIGNEUNIQUE', Num);
            TobDet.PutValue ('ACT_DATEMODIF', Nowh);
            if (TobDet.GetValue ('ACT_TYPEARTICLE') = 'PRE') then
              begin

              end
            else
            if (TobDet.GetValue ('ACT_TYPEARTICLE') = 'FRA') then
              begin

              end
            else
            if (TobDet.GetValue ('ACT_TYPEARTICLE') = 'MAR') then
              begin

              end;

            Inc (Num);
          end;

        // Enregistrement des modifs dans la base
        if not TobReg.InsertDB (nil) then
            V_PGI.IoError := oeUnknown;
      end;

  finally
    FiniMove;
    SourisNormale;
  End;
end;

procedure TOF_AFMODIFBONI.OnClose ;
begin

If EnErreur then begin LastError:=(-1); LastErrorMsg:=''; exit; end;

// Desallocations
TobReg.free;TobReg:=Nil;

  Inherited ;
end ;

procedure TOF_AFMODIFBONI.OnLoad ;
begin
  Inherited ;
EnErreur:=False;
bModified:=false;

end ;

procedure TOF_AFMODIFBONI.bFermerOnClick(SEnder: TObject);
begin
EnErreur := false;
  inherited;
end;

procedure TOF_AFMODIFBONI.bActuOnClick(SEnder: TObject);
Var mtt:double;
begin
  Mtt := TobReg.Detail[GS.Row-1].GetValue('ACT_TOTVENTE');
  Mtt := Mtt+ THNumEdit(GetControl('TOTDIFF')).value;
  TobReg.Detail[GS.Row-1].PutValue('ACT_TOTVENTE', Arrondi(Mtt, V_PGI.OkDecV));
  TobReg.Detail[GS.Row-1].PutValue('ACT_PUVENTE', Arrondi(Mtt, V_PGI.OkDecV));
    // Mise à jour du grid de saisie
  TobReg.PutGridDetail (GS, False, False, LesCol, true);
  bModified := true;
  PostMessage (GS.Handle, WM_KEYDOWN, VK_TAB, 0);
  application.processMessages;

end;

function AFLanceFiche_ModifBoni(Argument:String):string;
begin
Result:=AGLLanceFiche ('AFF','AFMODIFBONI','','',Argument);
end;


procedure AGLInsertionBoni(parms:array of variant; nb: integer ) ;
var  F : TForm ;
     MaTOF  : TOF;
begin
F:=TForm(Longint(Parms[0])) ;
if (F is TFvierge) then MaTOF:=TFvierge(F).LaTOF else exit;
if (MaTOF is TOF_AFMODIFBONI) then TOF_AFMODIFBONI(MaTOF).Insertion else exit;
end;


Initialization
  registerclasses ( [ TOF_AFMODIFBONI ] ) ;
  RegisterAglProc( 'InsertionBoni', True ,0, AGLInsertionBoni);
end.
