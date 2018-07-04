{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 13/05/2003
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : CPCODEAFB ()
Mots clefs ... : TOF;CPCODEAFB
*****************************************************************}

Unit CPCODEAFB_TOF;

//================================================================================
// Interface
//================================================================================
Interface

Uses
    StdCtrls, 
    Controls, 
    Classes, 
{$IFDEF EAGLCLIENT}
    MaineAGL,
{$ELSE}
    FE_Main,
    db, 
    dbtables, 
{$ENDIF}
    forms, 
    sysutils, 
    ComCtrls,
    HCtrls, 
    HEnt1, 
    HMsgBox, 
    UTOF,
    UTOB,
    HTB97
    ; 

//==================================================
// Externe
//==================================================
Procedure ParamCodeAFB;

//==================================================
// Definition de class
//==================================================
Type
    TOF_CPCODEAFB = Class (TOF)
        procedure OnNew                  ; override ;
        procedure OnDelete               ; override ;
        procedure OnUpdate               ; override ;
        procedure OnLoad                 ; override ;
        procedure OnArgument(S : String) ; override ;
        procedure OnDisplay              ; override ;
        procedure OnClose                ; override ;
        procedure OnCancel               ; override ;
    private
        FListe : THGrid;
        Combo : THValComboBox;
        Valider : TToolBarButton97;

        Query : TOB;
        Todel : TOB;

        OnCreat : boolean;

        svgNum,svgLib,svgCod : string;

        procedure OnClickMonValider(Sender : TObject);

        procedure OnRowEnterFliste(Sender : TObject ; ou : Longint ; var Cancel : Boolean ; Chg : Boolean);
        procedure OnRowExitFliste(Sender : TObject ; ou : Longint ; var Cancel : Boolean ; Chg : Boolean);

        procedure OnCellEnterFliste(Sender : TObject ; var ACol,ARow : Integer ; var Cancel : Boolean);
        procedure OnCellExitFliste(Sender : TObject ; var ACol,ARow : Integer ; var Cancel : Boolean);

        procedure OnChangeTTMODEPAIE(Sender : TObject);
        procedure OnChangeTTMODEPAIECODE(Sender : TObject);

        procedure SaveDB(row : integer);
        function  CanChange(col,row : integer ; var cancel : boolean) : boolean;
        function  HasChanged(num,lib,cod : string) : boolean;
    end;

//================================================================================
// Implementation
//================================================================================
Implementation

//==================================================
// Definition des Variables
//==================================================
var
    HMESS : array [0..1] of string = (
    {00}    '0;?CAPTION?;Vous devez renseigner un code.;W;O;O;O;',
    {01}    '0;?CAPTION?;Voulez vous abandonner la saisie ?;Q;YN;Y;N;'
            );

//==================================================
// fonctions hors class
//==================================================
{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 13/05/2003
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
Procedure ParamCodeAFB;
begin
    AGLLanceFiche('CP','CPCODEAFB','','','');
end;

//==================================================
// Evenements par default de la TOF
//==================================================
{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 13/05/2003
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPCODEAFB.OnNew;
var
    tmp : TOB;
begin
    Inherited;

    // combo existe ???
    if (not (combo = nil)) then FreeAndNil(combo);

    // en mose creation ??
    if (OnCreat = true) then exit;

    // Ajout d'une ligne dans la Grille
    if (Query.Detail.Count >= (FListe.RowCount-1)) then FListe.RowCount := FListe.RowCount+1
    else exit;

    // set de la variable de creation
    OnCreat := true;

    //tob fille sup !
    tmp := TOB.Create(Query.NomTable,Query,-1);
    tmp.InitValeurs;

    // Positionnement + evt
    FListe.gotoRow(FListe.RowCount-1);
    // Positionnement
    FListe.col := 1;
    FListe.SetFocus;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 13/05/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPCODEAFB.OnDelete;
var
    tmp : TOB;
    Cancel : boolean;
begin
    Inherited;

    // combo existe ???
    if (not (combo = nil)) then FreeAndNil(combo);

    // recupére la tob fille correspondante .....
    tmp := Query.Detail[FListe.Row-1];
    // l'effacé
    tmp.ChangeParent(ToDel,-1);
    // et updaté la base en effacant l'enreg correspondant ! (si on est pas en creation)
    if (not OnCreat) then ExecuteSQL('DELETE FROM CODEAFB WHERE AF_CODEAFB="' + FListe.Cells[1,FListe.Row] + '"')
    else OnCreat := false;
    // effacé la ligne du grid
    FListe.DeleteRow(FListe.Row);

    // on entre dans une nouvelle ligne
    OnRowEnterFliste(nil,0,Cancel,false);
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 13/05/2003
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPCODEAFB.OnUpdate;
var
    cancel : boolean;
begin
    Inherited;

    // combo existe ???
    if (not (combo = nil)) then FreeAndNil(combo);

    // maj
    if (CanChange(FListe.Col,FListe.Row,cancel)) then SaveDB(FListe.Row);

    // mettre a jour et PAS FERMER !
    // ???
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 13/05/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPCODEAFB.OnLoad;
var
    Cancel : boolean;
begin
    Inherited;

    // combo existe ???
    if (not (combo = nil)) then FreeAndNil(combo);

    OnRowEnterFliste(nil,0,Cancel,false);
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 13/05/2003
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... :
*****************************************************************}
procedure TOF_CPCODEAFB.OnArgument(S : String);
var
    BValider : TToolBarButton97;
begin
    Inherited;

    // Init des variable
    combo := nil;
    OnCreat := false;
    svgNum := '';
    svgLib := '';
    svgCod := '';

    // bouton valider
    Valider := TToolBarButton97(GetControl('MONVALIDER',true));
    BValider := TToolBarButton97(GetControl('BValider',true));

    Valider.OnClick := OnClickMonValider;
    Valider.Top := BValider.top;
    Valider.Left := BValider.left;

    // proprieté de la grille
    FListe := THGrid(GetControl('FListe',true));
    FListe.ColFormats[3] := 'CB=TTMODEPAIECODE||   ';
    FListe.ColFormats[4] := 'CB=TTMODEPAIE||   ';

    FListe.OnRowEnter := OnRowEnterFliste;
    FListe.OnRowExit := OnRowExitFliste;

    FListe.OnCellEnter := OnCellEnterFliste;
    FListe.OnCellExit := OnCellExitFliste;

    // creation du query
    Query := TOB.Create('CODEAFB',nil,-1);
    Query.LoadDetailDBFromSQL('CODEAFB','SELECT AF_CODEAFB, AF_LIBELLE, AF_MODEPAIEMENT LIBELLEPAIEMENT, AF_MODEPAIEMENT CODEPAIEMENT FROM CODEAFB');
    Query.PutGridDetail(FListe,false,false,'AF_CODEAFB;AF_LIBELLE;LIBELLEPAIEMENT;CODEPAIEMENT',true);
    // et de la liste des todel
    Todel := TOB.Create('TODEL',nil,-1);

    // resize des colones
    FListe.ColWidths[1] := Round((Fliste.ClientWidth-20) * 0.125);
    FListe.ColWidths[2] := Round((Fliste.ClientWidth-20) * 0.375);
    FListe.ColWidths[3] := Round((Fliste.ClientWidth-20) * 0.125);
    FListe.ColWidths[4] := Round((Fliste.ClientWidth-20) * 0.375);
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 13/05/2003
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... :
*****************************************************************}
procedure TOF_CPCODEAFB.OnClose;
var
    num,lib,cod : string;
begin
    num := FListe.Cells[1,FListe.row];
    lib := FListe.Cells[2,FListe.row];
    cod := FListe.Cells[3,FListe.row];

    Inherited;

    if (HasChanged(num,lib,cod) or OnCreat) then
    begin
        if (HShowMessage(HMESS[1],'','') = mrNo) then
        begin
            LastError := 1;
            exit;
        end;
    end;

    FreeAndNil(Combo);
    FreeAndNil(Query);
    FreeAndNil(ToDel);

    AvertirTable('TTAFB');
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 13/05/2003
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPCODEAFB.OnDisplay;
begin
    Inherited;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 13/05/2003
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPCODEAFB.OnCancel;
begin
    Inherited;

    // combo existe ???
    if (not (combo = nil)) then FreeAndNil(combo);

    // fair un cancel des changement ..... aie ....
    {$IFDEF EAGLCLIENT}
    {Ne compile pas en 2/3}
    Query.Seek(Fliste.Row-1);
    {$ENDIF}
    Query.PutLigneGrid(FListe,FListe.Row,false,false,'AF_CODEAFB;AF_LIBELLE;LIBELLEPAIEMENT;CODEPAIEMENT');
end;

//==================================================
// Autres Evenements
//==================================================
{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 14/05/2003
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPCODEAFB.OnClickMonValider(Sender : TObject);
var
    cancel : boolean;
begin
    // combo existe ???
    if (not (combo = nil)) then FreeAndNil(combo);

    // maj
    if (CanChange(FListe.Col,FListe.Row,cancel)) then SaveDB(FListe.Row);
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 13/05/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPCODEAFB.OnRowEnterFliste(Sender : TObject ; ou : Longint ; var Cancel : Boolean ; Chg : Boolean);
begin
    // sauvegarde des valeur pour comparaison !
    svgNum := FListe.Cells[1,FListe.Row];
    svgLib := FListe.Cells[2,FListe.Row];
    svgCod := FListe.Cells[3,FListe.Row];

    if ((not OnCreat) and (FListe.Col = 1)) then FListe.Col := 2;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 13/05/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPCODEAFB.OnRowExitFliste(Sender : TObject ; ou : Longint ; var Cancel : Boolean ; Chg : Boolean);
begin
    if (CanChange(FListe.Col,ou,cancel)) then SaveDB(ou);
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 13/05/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPCODEAFB.OnCellEnterFliste(Sender : TObject ; var ACol,ARow : Integer ; var Cancel : Boolean);
begin
    // premiere colonne ?? alors on entre pas ..
    if (FListe.Col = 1) then
    begin
        if (not OnCreat) then Cancel := true;
    end
    // sinon on set les combos
    else if (FListe.Col = 3) then
    begin
        combo := FListe.ValCombo;
        combo.OnChange := OnChangeTTMODEPAIECODE;
    end
    else if (FListe.Col = 4) then
    begin
        combo := FListe.ValCombo;
        combo.OnChange := OnChangeTTMODEPAIE;
    end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 13/05/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPCODEAFB.OnCellExitFliste(Sender : TObject ; var ACol,ARow : Integer ; var Cancel : Boolean);
begin
    if (CanChange(ACol,ARow,Cancel)) then if (not (combo = nil)) then FreeAndNil(combo);
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 13/05/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPCODEAFB.OnChangeTTMODEPAIE(Sender : TObject);
var
    temp : string;
begin
    if (Sender = nil) then exit;

    temp := THValComboBox(Sender).Value;
    FListe.Cells[3,FListe.Row] := temp;
    FListe.Cells[4,FListe.Row] := RechDom('TTMODEPAIE',temp,false);
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 13/05/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPCODEAFB.OnChangeTTMODEPAIECODE(Sender : TObject);
var
    temp : string;
begin
    if (Sender = nil) then exit;

    temp := THValComboBox(Sender).Value;
    Fliste.Cells[3,FListe.Row] := temp;
    FListe.Cells[4,FListe.Row] := RechDom('TTMODEPAIE',temp,false);
end;

//==================================================
// Autres fonctions de la class
//==================================================
{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 14/05/2003
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPCODEAFB.SaveDB(row : integer);
var
    num,lib,cod : string;
begin
    // get des valeur
    num := FListe.Cells[1,row];
    lib := FListe.Cells[2,row];
    cod := FListe.Cells[3,row];

    // si ca a changé ...
    if (HasChanged(num,lib,cod)) then
    begin
        if ((cod = '   ') or (cod = '')) then
        begin
            if (OnCreat) then ExecuteSQL('INSERT INTO CODEAFB (AF_CODEAFB,AF_LIBELLE) VALUES ("' + num + '","' + lib + '")')
            else ExecuteSQL('UPDATE CODEAFB SET AF_LIBELLE="' + lib + '", AF_MODEPAIEMENT="" WHERE AF_CODEAFB="' + num + '"');
        end
        else
        begin
            if (OnCreat) then ExecuteSQL('INSERT INTO CODEAFB VALUES ("' + num + '","' + lib + '","' + cod + '")')
            else ExecuteSQL('UPDATE CODEAFB SET AF_LIBELLE="' + lib + '", AF_MODEPAIEMENT="' + cod + '" WHERE AF_CODEAFB="' + num + '"');
        end;

        OnCreat := false;

        // mise a jour de la TOB
        {$IFDEF EAGLCLIENT}
        {Ne compile pas en 2/3}
        Query.Seek(Fliste.Row-1);
        {$ENDIF}
        Query.GetLigneGrid(FListe,row,'AF_CODEAFB;AF_LIBELLE;LIBELLEPAIEMENT;CODEPAIEMENT');
    end;

    // les nouvelle valeur sont les bonne
    svgNum := FListe.Cells[1,Row];
    svgLib := FListe.Cells[2,Row];
    svgCod := FListe.Cells[3,Row];
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 14/05/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
function TOF_CPCODEAFB.CanChange(col,row : integer ; var cancel : boolean) : boolean;
begin
    result := true;

    // premiere colone ?? alors pas si vide !!
    if (col = 1) then
    begin
        if (FListe.Cells[col,row] = '') then
        begin
            HShowMessage(HMESS[0],'','');
            Cancel := true;
            result := false;
        end;
    end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 14/05/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
function TOF_CPCODEAFB.HasChanged(num,lib,cod : string) : boolean;
begin
    result := false;

    // si ca a changé ...
    if (not ((svgNum = num) and (svgLib = lib) and (svgCod = cod))) then result := true;
end;

//================================================================================
// Initialization
//================================================================================
Initialization
    registerclasses([TOF_CPCODEAFB]);
end.
