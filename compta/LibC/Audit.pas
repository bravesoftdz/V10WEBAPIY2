{***********UNITE*************************************************
Auteur  ...... : BPY
Créé le ...... : 24/09/2004
Modifié le ... : 24/09/2004
Description .. : - BPY le 24/09/2004 - Passage en eAGL
Mots clefs ... : 
*****************************************************************}
unit Audit;

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
    StdCtrls,
    Buttons,
    ExtCtrls,
    hmsgbox,
    HSysMenu,
    Ent1,
    HEnt1,
{$IFNDEF EAGLCLIENT}
{$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
{$ENDIF}
    Hctrls,
    HTB97,
    HPanel,
    UiUtil,
    ParamSoc;

function LanceAuditPourCloture : Boolean;

type
    TFAudit = class(TForm)
        Panel1 : TPanel;
        HPB : TToolWindow97;
        BAide : TToolbarButton97;
        BValider : TToolbarButton97;
        BFerme : TToolbarButton97;
        V1 : THLabel;
        V2 : THLabel;
        V4 : THLabel;
        V5: THLabel;
        V6: THLabel;
        Panel2 : TPanel;
        R1 : THLabel;
        R2 : THLabel;
        R4 : THLabel;
        R5 : THLabel;
        R6 : THLabel;
        MsgRien : THMsgBox;
        MsgLibel : THMsgBox;
        HMTrad : THSystemMenu;
        Dock : TDock97;
        ToolbarButton971 : TToolbarButton97;
        B1 : TToolbarButton97;
        B2 : TToolbarButton97;
        B4 : TToolbarButton97;
        B5 : TToolbarButton97;
        B6 : TToolbarButton97;

        procedure FormShow(Sender : TObject);
        procedure BFermeClick(Sender : TObject);
        procedure BValiderClick(Sender : TObject);
        procedure BAideClick(Sender : TObject);
        procedure FormClose(Sender : TObject;var Action : TCloseAction);

        procedure B1Click(Sender : TObject);
        procedure B2Click(Sender : TObject);
        procedure B4Click(Sender : TObject);
        procedure B5Click(Sender : TObject);
        procedure B6Click(Sender : TObject);
    private
    { Déclarations privées }
        StR1,StR2,StR4,StR5,StR6 : string;
        OnSort : Boolean;
        Fait1,Fait2,Fait3,Fait4,Fait5,Fait6 : Boolean;

        procedure MajSociete;
        procedure ControlePerValid;
    public
    { Déclarations publiques }
    end;

implementation

{$R *.DFM}

uses
    VerCpta,
    VerCai,
    VerSolde,
    CloPerio,
    Cloture,
    Valperio;

{***********A.G.L.***********************************************
Auteur  ...... : 
Créé le ...... : 24/09/2004
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
function LanceAuditPourCloture : Boolean;
var
    Aud : TFAudit;
    OutProg : Boolean;
    PP : ThPanel; //SG6 05.04.05 FQ 14987
begin
    Result := FALSE;
    Aud := TFAudit.Create(Application);
    OutProg := FALSE;
    Aud.OnSort := OutProg;

    //SG6 05.04.05 FQ 14987
    PP := FindInsidePanel;

    if ((PP=nil)) then
    begin
        try
            Aud.ShowModal;
        Finally
            OutProg := Aud.OnSort;
            Aud.Free;
        end;
    end
    else
    begin
        InitInside(Aud,PP);
        Aud.Show;
    end;



    if OutProg then Result := true;
    Screen.Cursor := SyncrDefault;
end;

{***********A.G.L.***********************************************
Auteur  ...... : 
Créé le ...... : 24/09/2004
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TFAudit.FormShow(Sender : TObject);
var
{$IFDEF SPEC302}
    Q : TQuery;
{$ENDIF}
    StEtat, StRecup : string;
    i : Integer;
begin
    PopUpMenu := ADDMenuPop(PopUpMenu, '', '');
    StEtat := '';
    StRecup := '';
{$IFDEF SPEC302}
    Q := OpenSql(' Select SO_RECUPCPTA from SOCIETE where SO_SOCIETE="' + V_PGI.CodeSociete + '" ', True);
    if not Q.Eof then StRecup := Q.fields[0].AsString;
    Ferme(Q);
{$ELSE}
    StRecup := GetParamSoc('SO_RECUPCPTA');
{$ENDIF}
    for i := 1 to Length(StRecup) do
    begin
        if (StRecup[i] = '*') or (StRecup[i] = '-') or (StRecup[i] = '+') then
            StEtat := StEtat + StRecup[i]
        else
            StEtat := StEtat + '*';
    end;
    if (Length(StEtat) >= 6) then
    begin
        StR1 := StEtat[1];
        if StR1 = '-' then R1.Caption := MsgLibel.Mess[2]
        else if StR1 = '+' then R1.Caption := MsgLibel.Mess[1]
        else R1.Caption := MsgLibel.Mess[0];

        StR2 := StEtat[2];
        if StR2 = '-' then R2.Caption := MsgLibel.Mess[2]
        else if StR2 = '+' then R2.Caption := MsgLibel.Mess[1]
        else R2.Caption := MsgLibel.Mess[0];

        StR4 := StEtat[4];
        if StR4 = '-' then R4.Caption := MsgLibel.Mess[2]
        else if StR4 = '+' then R4.Caption := MsgLibel.Mess[1]
        else R4.Caption := MsgLibel.Mess[0];

        StR5 := StEtat[5];
        if StR5 = '-' then R5.Caption := MsgLibel.Mess[2]
        else if StR5 = '+' then R5.Caption := MsgLibel.Mess[1]
        else R5.Caption := MsgLibel.Mess[0];

        StR6 := StEtat[6];
        if StR6 = '-' then R6.Caption := MsgLibel.Mess[2]
        else if StR6 = '+' then R6.Caption := MsgLibel.Mess[1]
        else R6.Caption := MsgLibel.Mess[0];
    end
    else
    begin
        StR1 := '*'; R1.Caption := MsgLibel.Mess[0];
        StR2 := '*'; R2.Caption := MsgLibel.Mess[0];
        StR4 := '*'; R4.Caption := MsgLibel.Mess[0];
        StR5 := '*'; R5.Caption := MsgLibel.Mess[0];
        StR6 := '*'; R6.Caption := MsgLibel.Mess[0];
    end;

    Fait1 := False;
    Fait2 := False;
    Fait3 := False;
    Fait4 := False;
    Fait5 := False;
    Fait6 := False;

    if EstSerie(S5) then V5.Caption := MSGLibel.Mess[3];
end;

{***********A.G.L.***********************************************
Auteur  ...... : 
Créé le ...... : 24/09/2004
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TFAudit.BFermeClick(Sender : TObject);
begin
  Close ;
  //SG6 05.04.05 FQ 14987
  if IsInside(Self) then
  begin
    CloseInsidePanel(Self) ;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : 
Créé le ...... : 24/09/2004
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TFAudit.BValiderClick(Sender : TObject);
begin
    MajSociete;

// GP le 31/01/2001 if ClotureComptable(True) then OnSort:=TRUE ;

    if ctxPCL in V_PGI.PGIContexte then onSort := true
    else
    begin
        if ClotureComptable(True) then OnSort := true;
    end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : 
Créé le ...... : 24/09/2004
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TFAudit.B1Click(Sender : TObject);
begin
    if VerPourClo(VH^.EnCours) then
    begin
        R1.Caption := MsgLibel.Mess[1];
        StR1 := '+';
    end
    else
    begin
        R1.Caption := MsgLibel.Mess[2];
        StR1 := '-';
    end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : 
Créé le ...... : 24/09/2004
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TFAudit.B2Click(Sender : TObject);
begin
    ControlePerValid;
end;

{***********A.G.L.***********************************************
Auteur  ...... : 
Créé le ...... : 24/09/2004
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TFAudit.B4Click(Sender : TObject);
begin
    ControleSolde;
    R4.Caption := MsgLibel.Mess[1];
    StR4 := '+';
end;

{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... : 24/09/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TFAudit.B5Click(Sender : TObject);
begin
    case CtrlTo(VH^.EnCours) of
        cOk :
            begin
                R5.Caption := MsgLibel.Mess[1];
                StR5 := '+';
                Fait5 := True;
            end;
        cPasFait :
            begin
                if not Fait5 then
                begin
                    R5.Caption := MsgLibel.Mess[0];
                    StR5 := '*';
                end;
            end;
    end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : 
Créé le ...... : 24/09/2004
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TFAudit.B6Click(Sender : TObject);
begin
    case CtrlCaiClo of
        cPasFait :
            begin
                if not Fait6 then
                begin
                    R6.Caption := MsgLibel.Mess[0];
                    StR6 := '*';
                end;
            end;
        cOk :
            begin
                R6.Caption := MsgLibel.Mess[1];
                StR6 := '+';
                Fait6 := True;
            end;
        cPasOk :
            begin
                R6.Caption := MsgLibel.Mess[2];
                StR6 := '-';
                Fait6 := True;
            end;
    end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : 
Créé le ...... : 24/09/2004
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TFAudit.MajSociete;
var
    St : string;
{$IFDEF SPEC302}
    StSql : string;
{$ENDIF}
begin
    St := StR1 + StR2 + StR4 + StR5 + StR6;
{$IFDEF SPEC302}
    StSql := ' UPDATE SOCIETE Set SO_RECUPCPTA="' + St + '" where SO_SOCIETE="' + V_PGI.CodeSociete + '" ';
    ExecuteSql(StSql);
{$ELSE}
    SetParamSoc('SO_RECUPCPTA', St);
{$ENDIF}
end;

{***********A.G.L.***********************************************
Auteur  ...... : 
Créé le ...... : 24/09/2004
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TFAudit.ControlePerValid;
var
    Q : TQuery;
    i, NbPer : Byte;
    StValide : string;
    Valide : Boolean;
begin
    StValide := '';
    NbPer := VH^.Encours.NombrePeriode;
    Valide := False;
    Q := OpenSql('Select EX_VALIDEE from exercice where ex_exercice="' + VH^.Encours.Code + '" ', True);
    if not Q.Eof then StValide := Q.fields[0].AsString;
    Ferme(Q);
    if StValide = '' then Exit;
    for i := 1 to NbPer do
    begin
        if StValide[i] = '-' then
            Valide := False
        else
            Valide := True;
        if not Valide then Break;
    end;
    if Valide then
    begin
        R2.Caption := MsgLibel.Mess[1];
        StR2 := '+';
        MsgRien.Execute(2, caption, '');
    end
    else
    begin
        if (MsgRien.Execute(0, caption, '') <> mrYes) then Exit;
        case ValPerPourClo(True, False) of
            cPasFait :
                begin
                    if not Fait2 then
                    begin
                        R2.Caption := MsgLibel.Mess[0];
                        StR2 := '*';
                    end;
                end;
            cOk :
                begin
                    R2.Caption := MsgLibel.Mess[1];
                    StR2 := '+';
                    Fait2 := True;
                end;
            cPasOk :
                begin
                    R2.Caption := MsgLibel.Mess[2];
                    StR2 := '-';
                    Fait2 := True;
                end;
        end;
    end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : 
Créé le ...... : 24/09/2004
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TFAudit.BAideClick(Sender : TObject);
begin
    CallHelpTopic(Self);
end;

{***********A.G.L.***********************************************
Auteur  ...... : 
Créé le ...... : 24/09/2004
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TFAudit.FormClose(Sender : TObject;var Action : TCloseAction);
begin
    if Parent is THPanel then Action := caFree;
end;

end.
