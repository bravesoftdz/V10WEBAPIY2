{***********UNITE*************************************************
Auteur  ...... : JL
Créé le ...... : 11/08/2004
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : PGMULDECLACCTRAV ()
Mots clefs ... : TOF;PGMULDECLACCTRAV
*****************************************************************}
{
PT1   : 16/03/2007 V_80 RMA Déclaration accident du travail pour la MSA
}
Unit UTofPGMulDeclarations ;

Interface

Uses Classes,
{$IFNDEF EAGLCLIENT}
     db,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}mul,HDB,FE_Main,
{$else}
     eMul,uTob,MainEAgl,
{$ENDIF}
     sysutils,HCtrls,HEnt1,UTOF,HTB97 ,ParamDat,EntPaie,PgOutils2;

Type
  TOF_PGMULDECLACCTRAV = Class (TOF)
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    private
    Typedeclar : String;  //PT1

    procedure CreerDeclAcc(Sender : TOBject);
    procedure grilleDblClick(Sender : TObject);
    Procedure DateElipsisClick(Sender : TObject);
    Procedure ExitEdit(Sender : TObject);
  end ;

Implementation

procedure TOF_PGMULDECLACCTRAV.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_PGMULDECLACCTRAV.OnArgument (S : String ) ;
var BNew : TToolBarButton97;
    Q : TQuery;
    {$IFNDEF EAGLCLIENT}
    Liste : THDBGrid;
    {$ELSE}
    Liste : THGrid;
    {$ENDIF}
    Edit : THedit;
begin
  Inherited ;
  Typedeclar := Trim(S); //PT1

        Q := OpenSQL('SELECT * FROM EXERSOCIAL WHERE PEX_ACTIF="X" ORDER BY PEX_ANNEEREFER DESC', TRUE);
        if not Q.EOF then
        begin
                SetControlText('PDT_DATEACC',DateToStr(Q.FindField('PEX_DATEDEBUT').AsDateTime));
                SetControlText('PDT_DATEACC_',DateToStr(Q.FindField('PEX_DATEFIN').AsDateTime));
        end
        else
        begin
                SetControlText('PDT_DATEACC',DateToStr(Date));
                SetControlText('PDT_DATEACC_',DateToStr(Date));
        end;
        BNew  :=  TToolbarButton97 (getcontrol ('BInsert'));
        if BNew  <>  NIL then  BNew.OnClick  :=  CreerDeclAcc;
        {$IFNDEF EAGLCLIENT}
        Liste := THDBGrid(GetControl('FLISTE'));
        {$ELSE}
        Liste := THGrid(GetControl('FLISTE'));
        {$ENDIF}
        if Liste  <>  NIL then Liste.OnDblClick  :=  GrilleDblClick;
        Edit := THEdit(GetControl('PDT_SALARIE'));
        If Edit <> Nil then Edit.OnExit := ExitEdit;
        Edit := THEdit(GetControl('PDT_DATEACC'));
        If Edit <> Nil then Edit.OnElipsisClick := DateElipsisClick;
        Edit := THEdit(GetControl('PDT_DATEACC_'));
        If Edit <> Nil then Edit.OnElipsisClick := DateElipsisClick;
        //SetControlText('PDT_TYPEDECLAR','ACT');                   PT1
        SetControlText('PDT_TYPEDECLAR',Typedeclar);              //PT1
        If Typedeclar = 'MSA' Then                                //PT1
        Begin                                                     //PT1
           TFMul(Ecran).Caption := TFMul(Ecran).Caption + ' MSA'; //PT1
           UpdateCaption(TFMul(Ecran));                           //PT1
        End;                                                      //PT1
end ;

procedure TOF_PGMULDECLACCTRAV.ExitEdit(Sender: TObject);
var edit : thedit;
begin
edit:=THEdit(Sender);
if edit <> nil then	//AffectDefautCode que si gestion du code salarié en Numérique
    if (VH_Paie.PgTypeNumSal='NUM') and
       (length(Edit.text)<11) and
       (isnumeric(edit.text)) then
      edit.text:=AffectDefautCode(edit,10);
end;

procedure TOF_PGMULDECLACCTRAV.DateElipsisclick(Sender: TObject);
var
  key : char;
begin
  key := '*';
  ParamDate (Ecran, Sender, Key);
end;

procedure TOF_PGMULDECLACCTRAV.CreerDeclAcc(Sender : TOBject);
begin
   If Typedeclar = 'MSA' Then AGLLanceFiche('PAY','DECLACCTRAV_MSA','','','ACTION=CREATION;MSA')  //PT1
                         Else AGLLanceFiche('PAY','DECLACCTRAV','','','ACTION=CREATION;ACT');     //PT1
   TFMul(Ecran).BCherche.Click;
end;

procedure TOF_PGMULDECLACCTRAV.GrilleDblClick(Sender : TObject);
var St : String;
begin
        St := TFmul(Ecran).Q.FindField('PDT_SALARIE').AsString +';' +
        TFmul(Ecran).Q.FindField('PDT_TYPEDECLAR').AsString +';' +
        IntToStr(TFmul(Ecran).Q.FindField('PDT_ORDRE').AsInteger);
        If Typedeclar = 'MSA' Then AGLLanceFiche('PAY','DECLACCTRAV_MSA','',St,'ACTION=MODIFICATION;MSA') //PT1
                              Else AGLLanceFiche('PAY','DECLACCTRAV','',St,'ACTION=MODIFICATION;ACT');    //PT1
        TFMul(Ecran).BCherche.Click;
end;

Initialization
  registerclasses ( [ TOF_PGMULDECLACCTRAV ] ) ;
end.

