{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 06/12/2013
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BTSFAMTARIF ()
Mots clefs ... : TOF;BTSFAMTARIF
*****************************************************************}
Unit BTSFAMTARIF_TOF ;

Interface

Uses StdCtrls, 
     Controls, 
     Classes, 
{$IFNDEF EAGLCLIENT}
     db, 
     fe_main,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,
{$else}
     eMul,
     uTob,
     mainEagl,
{$ENDIF}
     forms,
     sysutils, 
     ComCtrls,
     HCtrls,
     HEnt1, 
     HMsgBox,
     Grids,
     Vierge,
     HTB97,
     UTOB,
     Lookup,
     UTOF ;

Type
  TOF_BTSFAMTARIF = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  private
    //
    FamilleTarif  : THEdit;
    SousFamArt    : THEdit;
    Libelle       : THEdit;
    //
    Action        : TActionFiche ;
    //
    GS            : THGrid;
    //
   	TOBGrille     : TOB;
    //
    fColNamesGS   : string;
    Falignement   : string;
    Ftitre        : string;
    fLargeur      : string;
    FF            : string;
    //
    BValider      : TToolbarButton97;
    BInsert       : TToolbarButton97;
    BDelete       : TToolbarButton97;
    //
    procedure AfficheLagrille(First: boolean=false);
    procedure CreateTOB;
    procedure ControleChamp(Champ, Valeur: String);
    procedure DefinieGrid;
    procedure DessineGrille(GS: THGrid);
    procedure DestroyTOB;
    procedure FamilleTarifOnExit(Sender : TObject);
    procedure GetObjects;
    procedure SetScreenEvents;
    procedure GSOnRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    procedure BDeleteOnClick(Sender: TObject);
    procedure BInsertOnClick(Sender: TObject);
    procedure BValiderOnClick(Sender: TObject);
    procedure FamilleTarifOnElipsisClick(Sender: TObject);
    //
  end ;

Implementation

procedure TOF_BTSFAMTARIF.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BTSFAMTARIF.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BTSFAMTARIF.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_BTSFAMTARIF.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_BTSFAMTARIF.OnArgument (S : String ) ;
var i       : Integer;
    Critere : string;
    Champ   : string;
    Valeur  : string;
begin
  Inherited ;

  //Détermination au début du programme du nombre de décimale
  FF:='#';
  if V_PGI.OkDecV>0 then
   begin
    FF:='# ##0.';
    for i:=1 to V_PGI.OkDecV-1 do
       begin
       FF:=FF+'0';
       end;
    FF:=FF+'0';
   end;

  //Chargement des zones ecran dans des zones programme
  GetObjects;

  CreateTOB;

  Critere := S;

  While (Critere <> '') do
  BEGIN
    i:=pos(':',Critere);
    if i = 0 then i:=pos('=',Critere);
    if i <> 0 then
       begin
       Champ:=copy(Critere,1,i-1);
       Valeur:=Copy (Critere,i+1,length(Critere)-i);
       end
    else
       Champ := Critere;
    Controlechamp(Champ, Valeur);
    Critere:=(Trim(ReadTokenSt(S)));
  END;

  //Gestion des évènement des zones écran
  SetScreenEvents;

  // définition de la grille
  DefinieGrid;

  //dessin de la grille
  DessineGrille(GS);

  FamilleTarif.MaxLength := 18;
  SousFamArt.MaxLength   := 18;
  Libelle.MaxLength      := 35;
  //
  FamilleTarif.text := '';
  FamilleTarif.SetFocus;
  //
  BValider.Visible  := False;
  BInsert.Visible   := False;
  BDelete.Visible   := False;

  Action := taModif;

end ;

procedure TOF_BTSFAMTARIF.OnClose ;
begin
  Inherited ;
  //
  DestroyTOB;
  //
  GS.ClearSelected;
  //

end ;

procedure TOF_BTSFAMTARIF.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BTSFAMTARIF.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_BTSFAMTARIF.GetObjects;
begin
  //
  GS           := THgrid(GetControl('Fliste'));
  //
  BValider     := TToolbarButton97(GetControl('BValider'));
  BInsert      := TToolbarButton97(GetControl('BInsert'));
  BDelete      := TToolbarButton97(GetControl('BDelete'));
  //
  FamilleTarif := THEdit (GetControl('BSF_FAMILLETARIF'));
  SousFamArt   := THEdit (GetControl('BSF_SOUSFAMTARART'));
  Libelle      := THEdit (GetControl('BSF_LIBELLE'));
  //
end;

procedure TOF_BTSFAMTARIF.CreateTOB;
begin

  TOBGrille   := TOB.Create('LA GRILLE', nil, -1);

end;

procedure TOF_BTSFAMTARIF.DestroyTOB;
begin

  FreeAndNil(TOBGrille);

end;

procedure TOF_BTSFAMTARIF.ControleChamp(Champ, Valeur: String);
begin

end;

procedure TOF_BTSFAMTARIF.SetScreenEvents;
begin

  BValider.OnCLick    := BValiderOnClick;
  BInsert.OnClick     := BInsertOnClick;
  BDelete.OnClick     := BDeleteOnClick;

  FamilleTarif.OnExit := FamilleTarifOnExit;
  FamilleTarif.OnElipsisClick := FamilleTarifOnElipsisClick;

  GS.OnRowEnter       := GSOnRowEnter;

end;

procedure TOF_BTSFAMTARIF.BValiderOnClick(Sender : TObject);
Var Indice  : Integer;
    TOBL    : TOB;
begin

  TOBL := Nil;

  if Action = taModif then Indice := GS.row;

  //Mise à jour de la TOB de Traitement des modif et création !
  If Action = TaCreat then
  begin
    TOBL := TOB.Create('BTSOUSFAMILLETARIF', TOBGRILLE, -1);
    TOBL.AddChampSupValeur('SEL','');
    TOBL.AddChampSupValeur('BSF_SOUSFAMTARART',SousFamArt.text);
    TOBL.AddChampSupValeur('BSF_LIBELLE',Libelle.text);
    TOBL.AddChampSupValeur('BSF_FAMILLETARIF',FamilleTarif.text);
  end
  else if Action = taModif then
  begin
    TOBL := TOBGRILLE.Detail[Indice -1];
    TOBL.PutValue('BSF_SOUSFAMTARART',SousFamArt.text);
    TOBL.PutValue('BSF_LIBELLE',Libelle.text);
    TOBL.PutValue('BSF_FAMILLETARIF',FamilleTarif.text);
  end;

  TOBL.SetAllModifie(True);
  TOBL.InsertOrUpdateDB(True);

  AfficheLagrille(False);

  BValider.Visible  := True;
  BInsert.Visible   := True;
  BDelete.Visible   := True;

  Action := TaModif;

end;

procedure TOF_BTSFAMTARIF.BInsertOnClick(Sender : TObject);
begin

  SousFamArt.text   := '';
  Libelle.text      := '';

  BValider.Visible  := True;
  BInsert.Visible   := False;
  BDelete.Visible   := False;

  Action            := TaCreat;

end;

procedure TOF_BTSFAMTARIF.BDeleteOnClick(Sender : TObject);
Var STSQL : String;
    Indice: Integer;
    TOBL : TOB;
begin

  if PGIAsk('Confirmez-vous la suppression ?') <> MrYes then exit;

  StSQl := 'DELETE BTSOUSFAMILLETARIF WHERE BSF_FAMILLETARIF="' + FamilleTarif.Text + '" AND BSF_SOUSFAMTARART="' + SousFamArt.text + '"';
  ExecuteSQL(STSQL);

  Indice := GS.row;

  //Mise à jour de la TOB de Traitement des modif et création !
  TOBL := TOBGRILLE.Detail[Indice -1];
  FreeAndNil(TOBL);

  AfficheLagrille(False);

end;

procedure TOF_BTSFAMTARIF.FamilleTarifOnElipsisClick(Sender : TObject);
var SavFamTarif : String;
begin

  SavFamTarif := FamilleTarif.Text;

  LookupCombo(FamilleTarif);

  if FamilleTarif.Text = '' then
  begin
    FamilleTarif.Text := savFamTarif;
    exit;
  end;

  FamilleTarifOnExit(Self);

end;

procedure TOF_BTSFAMTARIF.FamilleTarifOnExit(Sender : TObject);
Var STSQL : String;
    QQ    : TQuery;
Begin
  //
  BValider.Visible  := False;
  BDelete.Visible   := False;
  //
  TobGrille.ClearDetail;
  GS.VidePile(false);

  if FamilleTarif.text = '' then exit;

  //Contrôle existence famille tarif
  StSQL := 'SELECT BFT_FAMILLETARIF FROM BTFAMILLETARIF WHERE BFT_FAMILLETARIF="' + FamilleTarif.Text + '"';
  If not existeSql(Stsql) then
  begin
    PGIError('Famille Tarif ' + FamilleTarif.text + ' Inexistante', 'Famille Tarif');
    FamilleTarif.SetFocus;
    exit;
  end;

  //chargement de la tob qui servira à alimenter la grille
  StSQL := 'SELECT "" as SEL,BSF_SOUSFAMTARART, BSF_LIBELLE, BSF_FAMILLETARIF FROM BTSOUSFAMILLETARIF WHERE BSF_FAMILLETARIF="' + FamilleTarif.Text + '" ORDER BY BSF_SOUSFAMTARART';
  QQ    := OpenSQL(StSQL,False);

  IF Not QQ.EOF then
  begin
     TOBGrille.LoadDetailDB('BTSOUSFAMILLETARIF','','',QQ, False);
     AfficheLagrille(False);
     //
     BValider.Visible  := True;
     BDelete.Visible   := True;
     BInsert.Visible   := True;

     //
  end
  else
  begin
    SousFamArt.text := '';
    Libelle.text    := '';
    BValider.Visible  := True;
    BInsert.Visible   := False;
    BDelete.Visible   := False;
    Action := TaCreat;
  end;


end;


procedure TOF_BTSFAMTARIF.GSOnRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
begin

  //Affichage des zones écrans de la sous famille
  SousFamArt.text := GS.CellValues[1, GS.row];
  Libelle.text    := GS.CellValues[2, GS.row];

end;



procedure TOF_BTSFAMTARIF.DefinieGrid;
begin

  // Définition de la liste de saisie pour la grille Détail
  fColNamesGS := 'BSF_SOUSFAMTARART;BSF_LIBELLE';
  Falignement := 'C.0  ---;C.0  ---;';
  Ftitre := 'Sous-Famille;Libelle;';
  fLargeur := '20;80';

end;

procedure TOF_BTSFAMTARIF.DessineGrille(GS : THGrid);
var st              : String;
    lestitres       : String;
    lesalignements  : String;
    FF              : String;
    alignement      : String;
    Nam             : String;
    leslargeurs     : String;
    lalargeur       : String;
    letitre         : String;
    lelement        : string;
    //
    Obli            : Boolean;
    OkLib           : Boolean;
    OkVisu          : Boolean;
    OkNulle         : Boolean;
    OkCumul         : Boolean;
    Sep             : Boolean;
    Okimg           : boolean;
    //
    dec             : Integer;
    NbCols          : integer;
    indice          : Integer;
    Ind             : Integer;
begin

  //Calcul du nombre de Colonnes du Tableau en fonction des noms
  st := fColNamesGS;

  NbCols := 0;

  repeat
    lelement := READTOKENST (st);
    if lelement <> '' then
    begin
      inc(NbCols);
    end;
  until lelement = '';
  //
  GS.ColCount     := Nbcols+1;
  //
  st              := fColNamesGS ;
  lesalignements  := Falignement;
  lestitres       := Ftitre;
  leslargeurs     := fLargeur;

  //Mise en forme des colonnes
  for indice := 0 to Nbcols -1 do
  begin
    if GS.DBIndicator then
      Ind := Indice +1
    else
      Ind := Indice;
    //
    Nam := ReadTokenSt (St); // nom
    //
    alignement  := ReadTokenSt(lesalignements);
    lalargeur   := readtokenst(leslargeurs);
    letitre     := readtokenst(lestitres);
    //
    TransAlign(alignement,FF,Dec,Sep,Obli,OkLib,OkVisu,OkNulle,OkCumul) ;
    //
    GS.cells[Ind,0]     := leTitre;
    GS.ColNames [Ind]   := Nam;

    //Alignement des cellules
    if copy(Alignement,1,1)='G' then //Cadré à Gauche
      GS.ColAligns[ind] := taLeftJustify
    else if copy(Alignement,1,1)='D' then  //Cadré à Droite
      GS.ColAligns[ind] := taRightJustify
    else if copy(Alignement,1,1)='C' then
      GS.ColAligns[ind] := taCenter; //Cadré au centre

    //Colonne visible ou non
    if OkVisu then
  		GS.ColWidths[ind] := strtoint(lalargeur)*GS.Canvas.TextWidth('W')
    else
    	GS.ColWidths[ind] := -1;

    //Affichage d'une image ou du texte
    okImg := (copy(Alignement,8,1)='X');
    if (OkLib) or (okImg) then
    begin
    	GS.ColFormats[ind] := 'CB=' + Get_Join(Nam);
      if OkImg then
      begin
      	GS.ColDrawingModes[Ind]:= 'IMAGE';
      end;
    end;
    if (Dec<>0) or (Sep) then
    begin
    	if OkNulle then
        GS.ColFormats[ind] := FF+';; ;'
      else
      	GS.ColFormats[ind] := FF;
    end;
  end ;

end;


procedure TOF_BTSFAMTARIF.AfficheLagrille (First : boolean=false);
var indice : integer;
begin

  if TOBGrille.detail.count = 0 then
		GS.RowCount := TOBGrille.detail.count + 2
  else
		GS.RowCount := TOBGrille.detail.count + 1;

  TRY
    for Indice := 0 to TOBGrille.detail.count -1 do
    begin
      TOBGrille.detail[Indice].PutLigneGrid (GS,Indice+1,false,false,fColNamesGS);
    end;
  FINALLY
    GS.Refresh;
    GS.row := 1;
    //Affichage des zones écrans de la sous famille
    SousFamArt.text := GS.CellValues[1, GS.row];
    Libelle.text    := GS.CellValues[2, GS.row];
    GS.SetFocus;
  END;

  //TFVierge(ecran).HMTrad.ResizeGridColumns(GS);

end;


Initialization
  registerclasses ( [ TOF_BTSFAMTARIF ] ) ;
end.

