{***********UNITE*************************************************
Auteur  ...... : BPY
Créé le ...... : 27/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}

unit SAISVISU;

//================================================================================
// Interface
//================================================================================
interface

uses
    Windows,
    Messages,
    SysUtils,
    Classes,
    Graphics,
    Controls,
    Forms,
    Dialogs,
    SaisComm,
    Grids,
{$IFDEF EAGLCLIENT}
{$ELSE}
    DBGrids,
    HDB,
    DB,
    {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
{$ENDIF}

{$IFDEF VER150}
   Variants,
{$ENDIF}
    Hqry,
    SaisUtil,
    HEnt1,
    Ent1,
    HCtrls,
    ExtCtrls,
    StdCtrls,
    Buttons,
    Lettutil,
    HSysMenu,
    UTOB
    ;

{$IFDEF EAGLCLIENT}
    type tquery = TOB;
{$ENDIF}

//==================================================
// Externe
//==================================================
procedure VisuDernieresPieces(TPIECE : TLIST);

//==================================================
// Definition de class
//==================================================
type
    TFSaisVisu = class(TForm)
        PPied: TPanel;
    BValide: THBitBtn;
    BAbandon: THBitBtn;
    BAide: THBitBtn;
        H_DATEENTREE: TLabel;
        H_NBPIECES: TLabel;
        HMTrad: THSystemMenu;
    BZoomPiece: THBitBtn;
    GV: THGrid;

        procedure FormCreate(Sender : TObject);
        procedure FormShow(Sender : TObject);
        procedure FormClose(Sender : TObject; var Action : TCloseAction);

        procedure OnDblClickGV(Sender : TObject);

        procedure OnClickBAbandon(Sender : TObject);
        procedure OnClickBValide(Sender : TObject);
        procedure OnClickBAide(Sender : TObject);
    private
    public
        QV : TQuery;
        TPIECE : TList;
        Procedure RechercheEcritures;
    end;

//================================================================================
// Implementation
//================================================================================
implementation

Uses Saisie ;

{$R *.DFM}

//==================================================
// fonctions hors class : Point d'entré
//==================================================
{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 27/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure VisuDernieresPieces(TPIECE : TLIST);
var
    X : TFSaisVisu;
begin
    if (TPIECE.Count <= 0) then exit;
    X := TFSaisVisu.Create(Application);

    try
        X.TPIECE := TPIECE;
        X.ShowModal;
    finally
        X.Free;
    end;

    Screen.Cursor := SyncrDefault;
end;

//==================================================
// Fonction de la class : Evenement de la fiche
//==================================================
{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 27/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TFSaisVisu.FormCreate(Sender : TObject);
begin
    PopUpMenu := ADDMenuPop(PopUpMenu,'','');
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 27/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TFSaisVisu.FormShow(Sender : TObject);
begin
    H_DATEENTREE.Caption := Trim(H_DATEENTREE.Caption) + ' ' + DatetoStr(V_PGI.DateEntree);
    H_NBPIECES.Caption := Trim(H_NBPIECES.Caption) + ' ' + IntToStr(TPIECE.Count);

    ReChercheEcritures;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 27/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TFSaisVisu.FormClose(Sender : TObject; var Action : TCloseAction);
begin
    Ferme(QV);
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 27/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TFSaisVisu.OnDblClickGV(Sender : TObject);
{$IFDEF EAGLCLIENT}
var
    TOBListe,TOBLigne : TOB;
{$ELSE}
{$ENDIF}
begin
{$IFDEF EAGLCLIENT}
  	TOBListe := TOB.Create('Liste Ecriture',nil,-1);
    TOBLigne := QV.Detail[GV.row-1];
    TOB.Create('ECRITURE',TOBListe,-1);
    TOBListe.Detail[0].Dupliquer(TOBLigne,False,True);
    TrouveEtLanceSaisie(TOBListe,taConsult,'');

    FreeAndNil(TOBListe);
{$ELSE}
    QV.Locate('E_JOURNAL;E_DATECOMPTABLE;E_NUMEROPIECE',
              VarArrayOf([GV.Cells[1, GV.Row], GV.Cells[2, GV.Row], GV.Cells[3, GV.Row]]), []);
    TrouveEtLanceSaisie(QV,taConsult,'');
{$ENDIF}
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 27/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TFSaisVisu.OnClickBAbandon(Sender : TObject);
begin
    Close;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 27/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TFSaisVisu.OnClickBValide(Sender : TObject);
begin
    OnDblClickGV(Nil);
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 27/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TFSaisVisu.OnClickBAide(Sender : TObject);
begin
    CallHelpTopic(Self);
end;

//==================================================
// Fonction de la class : Autres fonctions
//==================================================
{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 27/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
Procedure TFSaisVisu.RechercheEcritures ;
Var i  : integer ;
    M  : RMVT ;
    St,StWhere : String ;
{$IFDEF EAGLCLIENT}
{$ELSE}
    QT : TOB;
{$ENDIF}
begin
    Ferme(QV);
    StWhere := 'SELECT E_JOURNAL,E_DATECOMPTABLE,E_NUMEROPIECE,E_GENERAL,E_AUXILIAIRE,E_REFINTERNE,E_DEBIT,E_CREDIT,E_MODESAISIE,E_NUMLIGNE,E_LIBELLE FROM ECRITURE WHERE ';

    for i := 0 to TPIECE.Count-1 do
    begin
        M := P_MV(TPIECE.Items[i]).R;
        St := '(' + WhereEcriture(tsGene,M,False) + ')';

        if (i = 0) then St := '(' + St                      // premiere
        else St := ' OR ' + St;                             // autre
        if (i = TPIECE.Count-1) then St := St + ')';        // derniere

        StWhere := StWhere+St ;
    end;
    StWhere := StWhere + ' AND E_NUMLIGNE=1' ;

    // recup des enreg
    QV := OpenSql(StWhere,false);

{$IFDEF EAGLCLIENT}
    QV.PutGridDetail(GV,false,true,'E_JOURNAL;E_DATECOMPTABLE;E_NUMEROPIECE;E_GENERAL;E_AUXILIAIRE;E_REFINTERNE;E_DEBIT;E_CREDIT;E_LIBELLE',true);
{$ELSE}
    QT := TOB.Create('',nil,-1);
    QT.LoadDetailDB('','','',QV,false);
    QT.PutGridDetail(GV,false,true,'E_JOURNAL;E_DATECOMPTABLE;E_NUMEROPIECE;E_GENERAL;E_AUXILIAIRE;E_REFINTERNE;E_DEBIT;E_CREDIT;E_LIBELLE',true);
    FreeAndNil(QT);
{$ENDIF}
end;

end.
