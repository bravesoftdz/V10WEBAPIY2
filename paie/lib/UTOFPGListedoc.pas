{***********UNITE*************************************************
Auteur  ...... : Paie PGI
Créé le ...... : 05/08/2003
Modifié le ... :   /  /
Description .. : Gestion des documents contrats de travail pour le
Suite ........ : publipostage Word
Mots clefs ... : PAIE;FUSIONWORD
*****************************************************************}
{
PT1 05/08/2003 V_42 SB FQ 10681 Exclusion des nouveaux formats de fichier
PT2 29/09/2004 V_50 JL FQ 11268 Exclusion médecine du travail et formation
}
unit UTOFPGListedoc;

interface
Uses StdCtrls,Controls,Classes,forms,sysutils,
{$IFNDEF EAGLCLIENT}

{$ENDIF}
     HTB97,HCtrls,HMsgBox,UTOF,Windows;
    
                             
Type
  TOF_PGLISTEDOC = Class (TOF)
      procedure OnArgument(Arguments : String ) ; override ;
      procedure OnUpdate; override ;
      procedure OnClose;  override;
      private
      Update: Boolean;
      procedure RecupDoc;
      procedure AjouterFichier(Sender: TObject);
      procedure DblClickListBox(Sender: TObject);

  END ;

implementation

uses UTOFPGSaisieSalRtf,EntPaie;

{ TOF_PGLISTEDOC }

procedure TOF_PGLISTEDOC.AjouterFichier(Sender: TObject);
Var
Nomfichier : string;
List : TListBox;
Edit : THEdit;
begin
Edit:=THEdit(getcontrol('FICHIER'));
if edit=nil then exit;
List := TListBox(GetControl('LISTDOC'));
if list=nil then exit;
if Edit.text<>'' then
   begin
   If PgiAsk('Voulez-vous copier ce fichier sous le répertoire : '+VH_Paie.PGCheminRech+'?',Ecran.caption)= mryes Then
     Begin
     NomFichier:=ExtractFileName(Edit.text);
     Copyfile(PChar(Edit.Text),PChar(VH_Paie.PGCheminRech+'\'+NomFichier),False);
     List.Items.Add(VH_Paie.PGCheminRech+'\'+NomFichier);
     List.ItemIndex:=List.Items.IndexOf(VH_Paie.PGCheminRech+'\'+NomFichier);
     Edit.Text:='';
     End
   Else
     Begin
     List.Items.Add(Edit.Text);
     List.ItemIndex:=List.Items.IndexOf(Edit.Text);
     Edit.Text:='';
     End;
   End
   Else
   PgiBox('Vous devez récupérer un fichier!','Rechercher un fichier');
end;

  
procedure TOF_PGLISTEDOC.OnArgument(Arguments: String);
Var List : TListBox;
Btn : TToolBarButton97;
begin
inherited;
Btn := TToolBarButton97(GetControl('BAjouter'));
if Btn<>nil then Btn.OnClick:=AjouterFichier;
RecupDoc;
Update:=False;
List:= TListBox(GetControl('LISTDOC'));
if List<>nil then List.OnDblClick:=DblClickListBox;
end;

procedure TOF_PGLISTEDOC.OnClose;
var  List : TListBox;
    Ipos : Integer;
begin
  inherited;
List := TListBox(GetControl('LISTDOC'));
if list=nil then exit;
Ipos:=List.ItemIndex;
if Ipos=-1 then PGPathDocContrat:=''
else
  if (List.Selected[ipos]=True) and (update=False) then
    Begin
    if PGIAsk('Voulez-vous abandonner le traitement?',Ecran.Caption)=mrYes then
      PGPathDocContrat:=''
    else
      PGPathDocContrat:=List.Items.Strings[List.ItemIndex];
    End;
end;

procedure TOF_PGLISTEDOC.OnUpdate;
var  List : TListBox;
    Ipos : Integer;
begin
  inherited;
Update:=True;
List := TListBox(GetControl('LISTDOC'));
if list=nil then exit;
Ipos:=List.ItemIndex;
PGPathDocContrat:=List.Items.Strings[IPos];
end;

procedure TOF_PGLISTEDOC.RecupDoc;
var
  sr: TSearchRec;
  FileAttrs: Integer;
  List : TListBox;
  stPath : string;
  Const ChaineEntier = ['0'..'9'] ;
begin
List := TListBox(GetControl('LISTDOC'));
if list=nil then exit;
FileAttrs := faAnyFile;
stPath := VH_Paie.PGCheminRech+'\*.DOC';
if FindFirst(stPath, FileAttrs, sr) = 0 then
  begin
  if (Sr.name<>'Certificat.doc') AND (Sr.name<>'Solde.doc') AND (Sr.name<>'Contrat.doc' )
  and (Sr.name<>'MedecineTravail.doc')  and (Sr.name<>'ConvocForm.doc') then  //PT2
     if (Sr.name[1] in ChaineEntier) then else //PT1
       List.Items.Add(VH_Paie.PGCheminRech+'\'+sr.Name);
  while (FindNext(sr) = 0) do
    if (Sr.name<>'Certificat.doc') AND (Sr.name<>'Solde.doc') AND (Sr.name<>'Contrat.doc' )
    and (Sr.name<>'MedecineTravail.doc')  and (Sr.name<>'ConvocForm.doc') then     //PT2
     if (Sr.name[1]  in ChaineEntier) then else //PT1
        List.Items.Add(VH_Paie.PGCheminRech+'\'+sr.Name);
  sysutils.FindClose(sr);
  end;
end;

procedure TOF_PGLISTEDOC.DblClickListBox(Sender: TObject);
begin
OnUpdate;
Ecran.Close;
end;

Initialization
registerclasses([TOF_PGLISTEDOC]);
end.

