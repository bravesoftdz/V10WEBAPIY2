{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 10/04/2003
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : CPMSGTVA ()
Mots clefs ... : TOF;CPMSGTVA
*****************************************************************}

Unit CPMSGTVA_TOF;

//================================================================================
// Interface
//================================================================================
Interface

Uses
    StdCtrls,
    Controls,
    Classes,
{$IFDEF EAGLCLIENT}
    MaineAGl,
{$ELSE}
    FE_Main,
    db,
    {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ENDIF}
    forms,
    sysutils,
    ComCtrls,
    HCtrls,
    HEnt1,
    HMsgBox,
    UTOF,
    extctrls
    ;

//==================================================
// Externe
//==================================================
function MsgTva(Direct : integer ; Deductible : Boolean ; D1,D2 : TDateTime ; RegimeTva,CodeTva : String) : Boolean;

//==================================================
// Definition de class
//==================================================
Type
    TOF_CPMSGTVA = Class (TOF)
    public
        procedure OnNew                  ; override ;
        procedure OnDelete               ; override ;
        procedure OnUpdate               ; override ;
        procedure OnLoad                 ; override ;
        procedure OnArgument(S : String) ; override ;
        procedure OnDisplay              ; override ;
        procedure OnClose                ; override ;
        procedure OnCancel               ; override ;
    private
    public
    private
        direct : integer;
    end;

//================================================================================
// Implementation
//================================================================================
Implementation

//==================================================
// Variable
//==================================================
var
    return : boolean;
    HMess : array [0..5] of string = (
    {00}    '1;?caption?;Confirmez-vous la prévalidation de l''édition des factures indirectes?;E;YN;Y;N',
    {01}    '0;?caption?;Confirmez-vous la prévalidation de l''édition des factures directes?;E;YN;Y;N',
    {02}    '2;?caption?;Confirmez-vous la prévalidation de l''édition des acomptes?;E;YN;Y;N',
    {03}    'Vous venez de procéder à l''édition liée au règlement "indirect" des factures.',
    {04}    'Vous venez de procéder à l''édition liée au règlement "direct" des factures.',
    {05}    'Vous venez de procéder à l''édition liée au règlement des acomptes non lettrés.'
            );


//==================================================
// fonctions hors class
//==================================================
function MsgTva(Direct : integer ; Deductible : Boolean ; D1,D2 : TDateTime ; RegimeTva,CodeTva : String) : Boolean;
begin
    return := false;

    if (Deductible) then AGLLanceFiche('CP','CPMSGTVA','','','ACTION=MODIFICATION;1;' + IntToStr(Direct) + ';' + DateToStr(D1) + ';' + DateToStr(D2) + ';' + RegimeTva + ';' + CodeTva)
    else AGLLanceFiche('CP','CPMSGTVA','','','ACTION=MODIFICATION;0;' + IntToStr(Direct) + ';' + DateToStr(D1) + ';' + DateToStr(D2) + ';' + RegimeTva + ';' + CodeTva);

    result := return;
end;

//==================================================
// Evenements par default de la TOF
//==================================================
{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 10/04/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPMSGTVA.OnNew;
begin
    Inherited;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 10/04/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPMSGTVA.OnDelete;
begin
    Inherited;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 10/04/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPMSGTVA.OnUpdate;
begin
    Inherited;

    if (HShowMessage(HMess[direct],'','') <> mrYes) then return := false
    else return := true;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 10/04/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPMSGTVA.OnLoad;
begin
    Inherited;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 10/04/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPMSGTVA.OnArgument(S : String);
var
    s1,s2 : string;
    deductible : integer;
    D1,D2 : String;
    Regime,Code : string;
    TYPETVA : TRadioGroup;
begin
    Inherited;

    S1 := UpperCase(S);
    S2 := ReadTokenSt(S1);                              // mode agl
    S2 := ReadTokenSt(S1); deductible := StrToInt(s2);  // deductible
    S2 := ReadTokenSt(S1); direct := StrToInt(S2);      // Direct ou pas
    S2 := ReadTokenSt(S1); D1 := S2;                    // date 1
    S2 := ReadTokenSt(S1); D2 := S2;                    // date 2
    S2 := ReadTokenSt(S1); Regime := S2;                // regime TVA
    S2 := ReadTokenSt(S1); Code := S2;                  // code TVA

    // Set Des Control
    SetControlText('FDATE1',D1);
    SetControlText('FDATE2',D2);
    if (Regime = '') then SetControlText('FREGIMETVA',Traduirememoire('<<Tous>>'))
    else SetControlText('FREGIMETVA',Regime);
    SetControlText('FCODETVA',RechDom('TTTVA',Code,false));

    // pour le radiogroup
    TYPETVA := TRadioGroup(GetControl('TYPETVA',true));
    if (not assigned(TYPETVA)) then close;
    TYPETVA.ItemIndex := deductible;

    // Autre
    SetControlText('HLabel1',HMess[direct+3]);
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 10/04/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPMSGTVA.OnClose;
begin
    Inherited;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 10/04/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPMSGTVA.OnDisplay;
begin
    Inherited;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 10/04/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPMSGTVA.OnCancel;
begin
    Inherited;

    return := false;
end;

//================================================================================
// Initialization
//================================================================================
Initialization
    registerclasses([TOF_CPMSGTVA]);
end.
