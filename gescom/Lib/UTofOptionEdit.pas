unit UTofOptionEdit;

interface

uses  M3FP, StdCtrls,Controls,Classes,forms,sysutils,ComCtrls,
      HCtrls,HEnt1,HMsgBox,UTOB,UTOF, AglInit, Agent,EntGC,
{$IFDEF EAGLCLIENT}
      MaineAGL,HPdfPrev,UtileAGL,
{$ELSE}
      {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}Fe_Main,
      EdtEtat,EdtDoc,
{$IFNDEF V530}
      EdtREtat,  EdtRDoc,
{$ENDIF}
{$ENDIF}
      HTB97;

Type
    TPrepaImp = function  : Integer of Object ;

Procedure EntreeOptionEdit(Tip,Nat: String; var Modele : String ; var Apercu,DeuxPages,First : boolean ;
                     Spages : TPageControl ; stSQL,Titre : String; PrepaImp : TPrepaImp);
Procedure EntreeOptionEditNP(Tip,Nat: String; var Modele : String ; var Apercu,DeuxPages,First : boolean ;
                     Spages : TPageControl ; stSQL,Titre : String; PrepaImp : TPrepaImp);

Type
    TOF_OptionEdit = Class (TOF)
    private
        TFETAT : THLabel ;
        FETAT : THValComboBox ;
        BPARAMETAT : TToolbarButton97;
        FAPERCU : TCHECKBOX ;
        FREDUIRE : TCHECKBOX ;
        // Initialisation Champs
        procedure InitChamps ;
        procedure ChargeListeEtat(First : Boolean) ;
        procedure BPARAMETATClick(Sender: TObject);
        procedure FAPERCUClick(Sender: TObject);
        procedure FREDUIREClick(Sender: TObject);
        procedure FETATChange(Sender: TObject);
    public
        procedure OnArgument (Arguments : String ) ; override;
        procedure OnUpdate ; override;

    END ;

Type RPARETAT = RECORD
                Tip,Nat,Modele,Titre,stSQL : String ;
                Apercu,DeuxPages,First : boolean ;
                Spages : TPageControl ;
                END ;

Var PEtat : RPARETAT;
    BeforeImp : TPrepaImp ;

implementation

{***********A.G.L.***********************************************
Auteur  ...... : Michel Richaud
Créé le ...... : 30/10/2000
Modifié le ... :   /  /    
Description .. : Lancement d'une édition de type Etat par l'intermédiare
Suite ........ : d'une fenêtre de paramètrage.
Mots clefs ... : EDITION
*****************************************************************}
Procedure EntreeOptionEdit(Tip,Nat: String; var Modele : String ; var Apercu,DeuxPages,First : boolean ;
                     Spages : TPageControl ; stSQL,Titre : String;  PrepaImp : TPrepaImp);
BEGIN
PEtat.Tip:=Tip; PEtat.Nat:=Nat; PEtat.Modele:=Modele; PEtat.Titre:=Titre;
PEtat.Apercu:=Apercu; PEtat.DeuxPages:=DeuxPages; PEtat.First:=First;
PEtat.Spages:=Spages;
PEtat.stSQL:=stSQL;
BeforeImp:=PrepaImp;
AglLanceFiche('GC','GCOPTIONEDIT','','','');
Modele:=PEtat.Modele;
Apercu:=PEtat.Apercu; DeuxPages:=PEtat.DeuxPages; First:=PEtat.First;
END;

Procedure EntreeOptionEditNP(Tip,Nat: String; var Modele : String ; var Apercu,DeuxPages,First : boolean ;
                     Spages : TPageControl ; stSQL,Titre : String;  PrepaImp : TPrepaImp);
BEGIN
PEtat.Tip:=Tip; PEtat.Nat:=Nat; PEtat.Modele:=Modele; PEtat.Titre:=Titre;
PEtat.Apercu:=Apercu; PEtat.DeuxPages:=DeuxPages; PEtat.First:=First;
PEtat.Spages:=Spages;
PEtat.stSQL:=stSQL;
BeforeImp:=PrepaImp;
AglLanceFiche('GC','GCOPTIONEDIT','','','NP');
Modele:=PEtat.Modele;
Apercu:=PEtat.Apercu; DeuxPages:=PEtat.DeuxPages; First:=PEtat.First;
END;
{==============================================================================================}
{================================= Evenement de la TOF ========================================}
{==============================================================================================}
Procedure TOF_OptionEdit.OnArgument (Arguments : String ) ;
BEGIN
TFETAT:=THLabel(GetControl('TFETAT')) ;
FETAT:=THValComboBox(GetControl('FETAT')) ;
FETAT.OnChange:=FETATChange;
BPARAMETAT:=TToolbarButton97(GetControl('BPARAMETAT')) ;
BPARAMETAT.OnClick:=BPARAMETATClick ;
if Arguments = 'NP' then
    begin
    TFETAT.Visible := False;
    FETAT.Visible := False;
    BPARAMETAT.Visible := False;
    end;
FAPERCU:=TCHECKBOX(GetControl('FAPERCU')) ;
FAPERCU.OnClick:=FAPERCUClick ;
FREDUIRE:=TCHECKBOX(GetControl('FREDUIRE')) ;
FREDUIRE.OnClick:=FREDUIREClick ;
InitChamps;
END;

procedure TOF_OptionEdit.OnUpdate;
BEGIN
inherited;
BeforeImp ;
LanceEtat(PEtat.Tip,PEtat.NAt,PEtat.Modele,PEtat.Apercu,False,PEtat.DeuxPages,PEtat.Spages,PEtat.stSQL,PEtat.Titre,False);
END;

{==============================================================================================}
{================================ Initialisation Champs =======================================}
{==============================================================================================}
procedure TOF_OptionEdit.InitChamps ;
BEGIN
FAPERCU.Checked:=PEtat.Apercu;
FREDUIRE.Checked:=PEtat.DeuxPages;
ChargeListeEtat(PEtat.First);
PEtat.First:=False;
END;

procedure TOF_OptionEdit.ChargeListeEtat(First : Boolean) ;
Var QQ : TQuery ;
    iDef,i_ind : Integer ;
    OldDef,StrSQL : String ;
BEGIN
Idef := 0;
if First then iDef:=-1 else OldDef:=FETAT.Value ;
   BEGIN
   FETAT.Items.Clear ; FETAT.Values.Clear ;
   StrSQL:='SELECT MO_CODE, MO_LIBELLE, MO_DEFAUT FROM MODELES WHERE MO_TYPE="'+PEtat.Tip+'" AND MO_NATURE="'+PEtat.Nat+'"' ;
   QQ:=OpenSQL(StrSQL,TRUE) ;
   While Not QQ.EOF do
      BEGIN
      i_ind:=FEtat.Items.Add(QQ.FindField('MO_LIBELLE').AsString) ;
      if QQ.FindField('MO_DEFAUT').AsString='X' then Idef:=i_ind ;
      FEtat.Values.Add(QQ.FindField('MO_CODE').AsString) ;
      QQ.Next
      END ;
   Ferme(QQ) ;
   if First then
      BEGIN
      if idef>0 then FEtat.ItemIndex:=idef else FEtat.ItemIndex:=FEtat.Values.IndexOf(PEtat.Modele) ;
      END else
      BEGIN
      if FEtat.Values.IndexOf(OldDef)>=0 then FEtat.Value:=OldDef else FEtat.ItemIndex:=0 ;
      END ;
   END ;
END;

{==============================================================================================}
{============================ Evenement lié aux Champs ========================================}
{==============================================================================================}
procedure TOF_OptionEdit.FAPERCUClick(Sender: TObject);
BEGIN
PEtat.Apercu:=FAPERCU.Checked;
END;

procedure TOF_OptionEdit.FREDUIREClick(Sender: TObject);
BEGIN
PEtat.DeuxPages:=FREDUIRE.Checked;
END;

procedure TOF_OptionEdit.FETATChange(Sender: TObject);
BEGIN
PEtat.Modele:=FETAT.Value;
END;

{==============================================================================================}
{============================ Evenement lié aux Boutons =======================================}
{==============================================================================================}
procedure TOF_OptionEdit.BPARAMETATClick(Sender: TObject);
BEGIN
{$IFDEF EAGLCLIENT}
// AFAIREEAGL ???????
{$ELSE}
if PEtat.Tip='E' then EditEtat(PEtat.Tip,PEtat.Nat,PEtat.Modele,False,PEtat.Spages, '', '')
                 else EditDocument(PEtat.Tip,PEtat.Nat,PEtat.Modele,False);
{$ENDIF}
END;

Initialization
registerclasses([TOF_OptionEdit]);

end.
 
