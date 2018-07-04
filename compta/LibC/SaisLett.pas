{***********UNITE*************************************************
Auteur  ...... : BPY
Créé le ...... : 17/03/2003
Modifié le ... :   /  /
Description .. : Porté en eAGL sous CPMULSAISLETT
Mots clefs ... : LETTRAGE;SAISIE
*****************************************************************}

unit SaisLett;

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
    Grids,
    StdCtrls,
    Buttons,
    ExtCtrls,
{$IFDEF EAGLCLIENT}
{$ELSE}
    DB,
    {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
    DBGrids,
{$ENDIF}
    Hqry,
    Mask,
    Hctrls,
    ComCtrls,
    HEnt1,
    Ent1,
    SaisUtil,
    LettUtil,
    hmsgbox,
    HSysMenu,
    UtilPGI,
    HTB97, ADODB
    ;

//==================================================
// Externe
//==================================================
procedure LettrerEnSaisie(X : RMVT ; NbEche : integer = 0 ; StWhere : String = '');
procedure LettrerEnSaisieTreso(X : RMVT ; Gen,Aux : String ; SansLesPartiel : Boolean);
procedure DelettrerEnSaisieTreso(X : RMVT ; Gen,Aux,CodeLettre : String ; SansLesPartiel : Boolean);

//==================================================
// Definition de class
//==================================================
type
    TFSaisLett = class(TForm)
        HP6: TPanel;
        BValider: TToolbarButton97;
        BAnnuler: TToolbarButton97;
        BAide: TToolbarButton97;
        Distinguer: TCheckBox;

        GL: TDBGrid;
        QEche: THQuery;
        SEche: TDataSource;
        Pages: TPageControl;

        TabSheet1: TTabSheet;
        E_ETATLETTRAGE: THCritMaskEdit;
        E_NUMECHE: THCritMaskEdit;
        E_ECHE: THCritMaskEdit;
        E_ETABLISSEMENT: THCritMaskEdit;
        E_NUMEROPIECE: THCritMaskEdit;
        E_DEVISE: THCritMaskEdit;
        E_NATUREPIECE: THCritMaskEdit;
        E_DATECOMPTABLE: THCritMaskEdit;
        E_JOURNAL: THCritMaskEdit;
        E_TRESOLETTRE: THCritMaskEdit;
        E_QUALIFPIECE: THCritMaskEdit;
        E_EXERCICE: THCritMaskEdit;

        HLS: THMsgBox;
        HMTrad: THSystemMenu;
        TT: TTimer;
        XX_WHERE: TEdit;

        procedure FormCreate(Sender: TObject);
        procedure FormShow(Sender: TObject);

        procedure BValiderClick(Sender: TObject);
        procedure BAnnulerClick(Sender: TObject);
        procedure BAideClick(Sender: TObject);

        procedure GLDblClick(Sender: TObject);

        procedure TTTimer(Sender: TObject);
    private
        StWhere : String;
    public
        X : RMVT;
        NbEche : integer;
    end;

//================================================================================
// Implementation
//================================================================================
implementation

uses
    Lettrage;

{$R *.DFM}

//==================================================
// fonctions hors class
//==================================================
{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 17/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure LettrerEnSaisie(X : RMVT ; NbEche : integer ; StWhere : String);
var
    FL : TFSaislett;
begin
    FL := TFSaisLett.Create(Application);

    try
        FL.X := X;
        FL.NbEche := NbEche;
        FL.StWhere := StWhere;
        FL.ShowModal;
    finally
        FL.Free;
    end;

    Screen.Cursor := SyncrDefault;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 17/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure LettrerEnSaisieTreso( X : RMVT ; Gen,Aux : String ; SansLesPartiel : Boolean) ;
var
    R : RLETTR;
begin
    if (not X.Treso) then exit;

    FillChar(R,Sizeof(R),#0) ;
    R.General := Gen;
    R.Auxiliaire := Aux;
    R.Appel := tlSaisieTreso;
    R.GL := Nil;
    R.CritMvt := ' AND E_GENERAL="'+R.General+'" AND E_AUXILIAIRE="'+R.Auxiliaire+'" AND E_ETATLETTRAGE<>"TL" AND E_ETATLETTRAGE<>"RI" ';
    R.CritDev := X.CodeD;
    R.LettrageDevise := (X.CodeD <> V_PGI.DevisePivot);
    R.DeviseMvt := X.CodeD;
    R.Ident := X;
    R.Ident.NumLigne := X.NumLigne;
    R.Ident.NumEche := X.NumEche;
    R.SansLesPartiel := SansLesPartiel;
    LettrageManuel(R,True,taModif);
end ;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 17/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure DelettrerEnSaisieTreso( X : RMVT ; Gen,Aux,CodeLettre : String ; SansLesPartiel : Boolean) ;
var
    R : RLETTR;
begin
    if (not X.Treso) then exit;

    FillChar(R,Sizeof(R),#0);
    R.General := Gen;
    R.Auxiliaire := Aux;
    R.Appel := tlSaisieTreso;
    R.GL := Nil;
    R.CritMvt := ' AND E_GENERAL="'+R.General+'" AND E_AUXILIAIRE="'+R.Auxiliaire+'" AND E_ETATLETTRAGE<>"RI" ';
    R.CritDev := X.CodeD;
    R.LettrageDevise := (X.CodeD <> V_PGI.DevisePivot);
    R.DeviseMvt := X.CodeD;
    R.Ident := X;
    R.Ident.NumLigne := X.NumLigne;
    R.Ident.NumEche := X.NumEche;
    R.SansLesPartiel := SansLesPartiel;
    R.CodeLettre := CodeLettre;

    LettrageManuel(R,False,taModif);
end;


//==================================================
// Evenements de la fiche
//==================================================
{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 17/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TFSaisLett.FormCreate(Sender: TObject);
begin
    PopUpMenu := ADDMenuPop(PopUpMenu,'','');
    QEche.Manuel := True;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 17/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TFSaisLett.BAnnulerClick(Sender: TObject);
begin
    Close;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 17/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TFSaisLett.FormShow(Sender: TObject);
begin
    E_JOURNAL.Text       := X.Jal;
    E_DATECOMPTABLE.Text := DateToStr(X.DateC);
    E_NATUREPIECE.Text   := X.Nature;
    E_DEVISE.Text        := X.CodeD;
    E_NUMEROPIECE.Text   := IntToStr(X.Num);
    E_ETABLISSEMENT.Text := X.Etabl;
    E_QUALIFPIECE.Text   := X.Simul;
    E_EXERCICE.Text      := X.Exo;

    XX_WHERE.Text := StWhere;

    QEche.Liste := 'SAISLETT';
    QEche.Manuel := False;
    QEche.UpdateCriteres;

    CentreDBGrid(GL);

    if (VH^.TenueEuro) then Distinguer.Caption := HLS.Mess[1]+' '+RechDom('TTDEVISETOUTES',V_PGI.DevisePivot,False)+' + '+RechDom('TTDEVISETOUTES',V_PGI.DeviseFongible,False);
    if (X.CodeD <> V_PGI.DevisePivot) then Distinguer.Visible := False;
    if (NbEche = 1) then TT.Enabled := true;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 17/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TFSaisLett.BValiderClick(Sender: TObject);
begin
    GLDblClick(Nil);
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 17/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TFSaisLett.GLDblClick(Sender: TObject);
var
    R : RLETTR;
    CodeL : String4;
begin
    if (QEche.EOF) then exit;

    FillChar(R,Sizeof(R),#0);
    CodeL := GL.Fields[6].AsString;

    if ((CodeL<>'') and (CodeL=uppercase(CodeL))) then
    begin
        HLS.Execute(0,caption,'');
        exit;
    end;

    R.General := GL.Fields[0].AsString;
    R.Auxiliaire := GL.Fields[1].AsString;
    R.Appel := tlSaisieCour;
    R.GL := Nil;
    R.CritMvt := ' AND E_GENERAL="'+R.General+'" AND E_AUXILIAIRE="'+R.Auxiliaire+'" AND E_ETATLETTRAGE<>"TL" AND E_ETATLETTRAGE<>"RI" ';
    R.CritDev := X.CodeD;
    R.DeviseMvt := X.CodeD;
    R.LettrageDevise := (X.CodeD <> V_PGI.DevisePivot);
    R.Ident := X;
    R.Ident.NumLigne := GL.Fields[7].AsInteger;
    R.Ident.NumEche := GL.Fields[8].AsInteger;
    R.Distinguer := Distinguer.Checked;
    if (R.DeviseMvt = V_PGI.DevisePivot) then
    begin
        {Paquet en Franc ou Euro}
        if (R.Distinguer) then
            R.CritDev := V_PGI.DevisePivot;
    end ;
LettrageManuel(R,True,taModif) ;
QEche.UpdateCriteres ;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 17/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TFSaisLett.BAideClick(Sender: TObject);
begin
    CallHelpTopic(Self);
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 17/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TFSaisLett.TTTimer(Sender: TObject);
begin
    TT.Enabled := False;
    TT.Interval := 0;
    GLDblClick(Nil);
end;

end.
