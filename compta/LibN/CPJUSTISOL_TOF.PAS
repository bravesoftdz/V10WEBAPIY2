{***********UNITE*************************************************
Auteur  ...... : Benoit Python
Créé le ...... : 19/05/2003
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : CPJUSTISOL ()
Mots clefs ... : TOF;CPJUSTISOL
*****************************************************************}

Unit CPJUSTISOL_TOF;

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
    eMul,
{$ELSE}
    FE_Main,
    Mul,
    HDB,
    db,
    {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
{$ENDIF}

{$IFDEF COMPTA}
    Saisie,
{$ENDIF}
    forms,
    sysutils,
    ComCtrls,
    HCtrls,
    Ent1,
    HMsgBox,
    UTOF,
    SAISUTIL,
    UTOB,
    {$IFDEF MODENT1}
    CPTypeCons,
    {$ENDIF MODENT1}
    HQRY
    ;

//==================================================
// Externe
//==================================================
procedure JustifSolde(LeCpte : String ; Lefb : TFichierBase);

//==================================================
// Definition de class
//==================================================
Type
    TOF_CPJUSTISOL = Class(TOF)
        procedure OnNew                  ; override ;
        procedure OnDelete               ; override ;
        procedure OnUpdate               ; override ;
        procedure OnLoad                 ; override ;
        procedure OnArgument(S : String) ; override ;
        procedure OnDisplay              ; override ;
        procedure OnClose                ; override ;
        procedure OnCancel               ; override ;
    private
        cpt : String;
{$IFDEF EAGLCLIENT}
        FListe : THGrid;
{$ELSE}
        FListe : THDBGrid;
{$ENDIF}
        E_EXERCICE : THValComboBox;

        procedure OnChangeE_EXERCICE(Sender : TObject);
        procedure OnDblClickFListe(Sender : TObject);

        procedure LibelleAuxi;
        procedure CalculLeSolde;
    end;

//================================================================================
// Implementation
//================================================================================
Implementation

uses
  {$IFDEF MODENT1}
  ULibExercice,
  CPProcGen,
  CPProcMetier,
  {$ENDIF MODENT1}
  HEnt1;


//==================================================
// Variable local
//==================================================
var
    TypeFile : TFichierBase;

//==================================================
// Constant
//==================================================
const
    MESS : array [0..1] of string = (
        {00}    'Détail des échéances du compte Général : ',
        {01}    'Détail des échéances du compte de Tiers : '
        );

//==================================================
// Externe
//==================================================
procedure JustifSolde(LeCpte : String ; Lefb : TFichierBase);
begin
    TypeFile := Lefb;
    AglLanceFiche('CP','CPJUSTISOL','','','ACTION=CONSULTATION;' + LeCpte);
end;

//==================================================
// Evenements par default de la TOM
//==================================================
{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 19/05/2003
Modifié le ... :   /  /
Description .. : 
Mots clefs ... :
*****************************************************************}
procedure TOF_CPJUSTISOL.OnNew;
begin
    Inherited;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 19/05/2003
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPJUSTISOL.OnDelete;
begin
    Inherited;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 19/05/2003
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPJUSTISOL.OnUpdate;
begin
    Inherited;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 19/05/2003
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPJUSTISOL.OnLoad;
begin
    Inherited;

    CalculLeSolde;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 19/05/2003
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPJUSTISOL.OnArgument(S : String);
var
    S1,S2 : String;
    D1,D2 : TDateTime;
begin
    Inherited;

    // recup des parametre
    S1 := UpperCase(S);
    S2 := ReadTokenSt(S1);              // type action
    S2 := ReadTokenSt(S1); cpt := S2;   // compte

    // set des date
    SetControlText('E_DATECOMPTABLE',StDate1900);
    SetControlText('E_DATECOMPTABLE_',StDate2099);
    SetControlText('E_DATEECHEANCE',StDate1900);
    SetControlText('E_DATEECHEANCE_',StDate2099);

    // set du num compte
    Case (TypeFile) of
        fbAux :
            begin
                SetControlText('E_AUXILIAIRE',cpt);
                SetControlText('E_GENERAL','');
                TFMul(Ecran).DBListe := 'CPJUSTIFSLDA';
                Ecran.Caption := MESS[1] + cpt;
            end;
        fbGene :
            begin
                SetControlText('E_AUXILIAIRE','');
                SetControlText('E_GENERAL',cpt);
                TFMul(Ecran).DBListe := 'CPJUSTIFSLDG';
                Ecran.Caption := MESS[0] + cpt ;
            end;
    end;

    // recup de la liste
{$IFDEF EAGLCLIENT}
    FListe := THGrid(GetControl('FListe',true));
{$ELSE}
    FListe := THDBGrid(GetControl('FListe',true));
{$ENDIF}
    FListe.OnDblClick := OnDblClickFListe;

    // recup de l'exo
    E_EXERCICE := THValComboBox(GetControl('E_EXERCICE',true));
    E_EXERCICE.OnChange := OnChangeE_EXERCICE;
    E_EXERCICE.Value := QuelDateExo(V_PGI.DateEntree,D1,D2);
    ExoToDates(E_EXERCICE.Value,GetControl('E_DATECOMPTABLE',true),GetControl('E_DATECOMPTABLE_',true));
    // set des valeur par default
    SetControlText('E_DEVISE',V_PGI.DevisePivot);
    THValComboBox(GetControl('E_ETABLISSEMENT',true)).ItemIndex := 0;
    THValComboBox(GetControl('E_ETATLETTRAGE',true)).ItemIndex := 0;
    THValComboBox(GetControl('E_QUALIFPIECE',true)).ItemIndex := 0;
    // libelle auxiliaire
    LibelleAuxi;
    // mask d'affichage
    ChangeMask(THNumEdit(GetControl('CPTTOTDEBIT',true)),V_PGI.OkDecV,V_PGI.SymbolePivot);
    ChangeMask(THNumEdit(GetControl('CPTTOTCREDIT',true)),V_PGI.OkDecV,V_PGI.SymbolePivot);
    // caption
    UpdateCaption(Ecran);
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 19/05/2003
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPJUSTISOL.OnClose;
begin
    Inherited;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 19/05/2003
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPJUSTISOL.OnDisplay;
begin
    Inherited;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 19/05/2003
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPJUSTISOL.OnCancel;
begin
    Inherited;
end;

//==================================================
// Autres Evenements
//==================================================
{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 19/05/2003
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPJUSTISOL.OnChangeE_EXERCICE(Sender : TObject);
begin
    ExoToDates(E_EXERCICE.Value,GetControl('E_DATECOMPTABLE',true),GetControl('E_DATECOMPTABLE_',true));
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 19/05/2003
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPJUSTISOL.OnDblClickFListe(Sender: TObject);
begin
{$IFNDEF IMP}
{$IFDEF EAGLCLIENT}
    TFMul(Ecran).Q.TQ.Seek(FListe.Row-1);
    TrouveEtLanceSaisie(TFMul(Ecran).Q.TQ,taConsult,'');
{$ELSE}
    TrouveEtLanceSaisie(TFMul(Ecran).Q,taConsult,'');
{$ENDIF}
{$ENDIF}
end;

//==================================================
// Autres fonctions de la class
//==================================================
{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 19/05/2003
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPJUSTISOL.LibelleAuxi;
var
    Q1 : TQuery;
    St : String;
begin
    st := '';
    
//    if (Q.Eof) then
//    begin
//        TT_LIBELLE.Caption := '';
//        exit;
//    end;
    
    Case (TypeFile) of
        fbAux :
            begin
                Q1 := OpenSql('SELECT T_LIBELLE FROM TIERS WHERE T_AUXILIAIRE="' + cpt + '"',true);
                St := Q1.Fields[0].AsString;
                ferme(Q1);
            end;
        fbGene :
            begin
                Q1 := OpenSql('SELECT G_LIBELLE FROM GENERAUX WHERE G_GENERAL="' + cpt + '"',true);
                st := Q1.Fields[0].AsString;
                ferme(Q1);
            end;
    end;

    SetControlText('CPTNAME',St);
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 19/05/2003
Modifié le ... :   /  /    
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPJUSTISOL.CalculLeSolde;
var
    Q1 : TQuery;
    debit,credit : double;
begin
    Q1 := OpenSQL('SELECT SUM(E_CREDIT),SUM(E_DEBIT) FROM ECRITURE ' + RecupWhereCritere(TPageControl(GetControl('Pages',true))) + 'GROUP BY E_AUXILIAIRE',true);
    debit := Q1.Fields[0].AsFloat;
    credit := Q1.Fields[1].AsFloat;
    ferme(Q1);

    THNumEdit(GetControl('CPTTOTDEBIT',true)).Value := credit;
    THNumEdit(GetControl('CPTTOTCREDIT',true)).Value := debit;
    AfficheLeSolde(THNumEdit(GetControl('CPTTOTSOLDE',true)),credit,debit)
end;

//================================================================================
// Initialization
//================================================================================
Initialization
    registerclasses([TOF_CPJUSTISOL]);
end.
