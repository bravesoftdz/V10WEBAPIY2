{***********UNITE*************************************************
Auteur  ...... : Stéphane BOUSSERT
Créé le ...... : 21/03/2002
Modifié le ... : 26/04/2002
Description .. : Source TOF de la FICHE : CPNATCPTECODELG ()
Mots clefs ... : TOF;CPNATCPTECODELG
*****************************************************************}

Unit CPNATCPTECODELG_TOF;

Interface

Uses
    StdCtrls,
    Controls,
    Classes,
    forms,
    sysutils,
    ComCtrls,
    HCtrls,
    HEnt1,
    HMsgBox,
    UTOF,
    Spin,
    utilPGI,
{$IFDEF EAGLCLIENT}
    maineagl,
{$ELSE}
    FE_main,
{$ENDIF}
    ENT1
    ;

procedure CPLanceFiche_CodeLongueur(CodeTable : String);

Type
    TOF_CPNATCPTECODELG = Class (TOF)
        procedure OnUpdate               ; override ;
        procedure OnArgument(S : String) ; override ;
    private
        CodeTableLibre : String;
        seCodeLongueur : TSpinEdit;

        procedure CodeLongueurChange(Sender : TObject);
        procedure InitCodeLongueur;
    end;

Implementation

//==============================================================================
procedure CPLanceFiche_CodeLongueur(CodeTable : String);
begin
    AGLLanceFiche('CP','CPNATCPTECODELG','','',CodeTable);
end;

//==============================================================================
procedure TOF_CPNATCPTECODELG.OnUpdate ;
Var
	size,i,j : Integer ;
        lStDossier : string ;
begin
    Inherited;

    if (seCodeLongueur.Text = '') then
    begin
        PGIBox(TraduireMemoire('La longueur du code doit être comprise entre 1 et 17'),TraduireMemoire('Longueur du code'));
        SetFocusControl('CodeLongueur');
        InitCodeLongueur;
        LastError := 1;
        exit;
    end;

    size := seCodeLongueur.value;

    // Gestion du multi-dossier
    if EstTablePartagee( 'NATCPTE' )
      then lStDossier := TableToBase( 'NATCPTE' )
      else lStDossier := '' ;

    // Update de la table COMMUN
    ExecuteSQL('UPDATE ' + GetTableDossier( lStDossier, 'CHOIXCOD') + ' SET CC_LIBRE="' + IntToStr(size) + '" ' + 'WHERE CC_TYPE="NAT" and CC_CODE="' + CodeTableLibre + '"');

    // MAJ de la variable global VH
    Case CodeTableLibre[1] Of
        'G' : i := 1;
        'T' : i := 2;
        'S' : i := 3;
        'B' : i := 4;
        'D' : i := 5;
        'E' : i := 6;
        'A' : i := 7;
        'U' : i := 8;
        'I' : i := 9;
    else i := 0;
    end;

    if i > 0 Then
    begin
        j := StrToInt(Copy(CodeTableLibre,2,2)) + 1;
        VH^.LgTableLibre[i,j] := size;
    end;
end;

procedure TOF_CPNATCPTECODELG.OnArgument (S : String ) ;
begin
    Inherited;

    CodeTableLibre := S;
    seCodeLongueur := TSpinEdit(GetControl('CodeLongueur'));
    InitCodeLongueur;
    seCodeLongueur.OnChange := CodeLongueurChange;
end;

procedure TOF_CPNATCPTECODELG.CodeLongueurChange(Sender: TObject);
Var
    size : Integer;
begin
    if (seCodeLongueur.Text <> '') then
    begin
        size := seCodeLongueur.value;
        if (size < 1) or (size > 17) then
        Begin
            PGIBox(TraduireMemoire('La longueur du code doit être comprise entre 1 et 17'),TraduireMemoire('Longueur du code'));
            SetFocusControl('CodeLongueur');
            InitCodeLongueur;
        End;
    end;
end;

procedure TOF_CPNATCPTECODELG.InitCodeLongueur;
var
    i,j : Integer;
begin
    Case CodeTableLibre[1] Of
        'G' : i := 1;
        'T' : i := 2;
        'S' : i := 3;
        'B' : i := 4;
        'D' : i := 5;
        'E' : i := 6;
        'A' : i := 7;
        'U' : i := 8;
        'I' : i := 9;
    else i := 0;
    end;

    if i>0 Then
    begin
        j := StrToInt(Copy(CodeTableLibre,2,2))+1 ;
        seCodeLongueur.value := VH^.LgTableLibre[i,j] ;
    end;
end;

Initialization
    registerclasses([TOF_CPNATCPTECODELG]);
end.
