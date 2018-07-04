unit DpTableauBordDetailSolde;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Vierge, Grids, Hctrls, ExtCtrls, HSysMenu, HTB97, Utob,
{$IFDEF EAGLCLIENT}
{$ELSE}
  {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
{$ENDIF}
  Hent1,
  HmsgBox, UTofConsEcr, DpTableauBordLibrairie;

type
  TFDPTableauBordDetailSolde = class(TFVierge)
    Panel1: TPanel;
    GrilleDetailSolde: THGrid;
    procedure FormShow(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure GrilleDetailSoldeDblClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure GrilleDetailSoldeKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);

  private
   NumDossier : String;
   TobSolde   : Tob;

   procedure InitialiserGrilleDetailSolde ();
   procedure AjouterElementGrilleSolde (Var UneTob : Tob; ChDossier,ChLibelle : String);
   procedure AjouterTotalGrilleSolde (Var UneTob : Tob);
   procedure GetCellCanvas(ACol,ARow: LongInt;  Canvas : TCanvas; AState : TGridDrawState);
   procedure PostDrawCell(ACol,ARow: Integer; Canvas: TCanvas; AState: TGridDrawState);

  end;

//------------------------------------
//--- Déclaration procedure globale
//------------------------------------
procedure AfficherDetailSoldeClient (NumDossier : String);

implementation

{$R *.DFM}

//----------------------------------------
//--- Nom   : AfficherDetailSoldeClient
//--- Objet :
//----------------------------------------
procedure AfficherDetailSoldeClient (NumDossier : String);
var FDPTableauBordDetailSolde: TFDPTableauBordDetailSolde;
begin
 FDPTableauBordDetailSolde:=TFDPTableauBordDetailSolde.Create(Application);
 FDPTableauBordDetailSolde.NumDossier:=NumDossier;
 try
  FDPTableauBordDetailSolde.ShowModal;
 finally
  FDPTableauBordDetailSolde.Free;
 end;
end;

//----------------------------------------
//--- Nom   : FormCreate
//--- Objet :
//----------------------------------------
procedure TFDPTableauBordDetailSolde.FormCreate(Sender: TObject);
begin
 inherited;
 GrilleDetailSolde.GetCellCanvas:=GetCellCanvas ;
 GrilleDetailSolde.PostDrawCell:=PostDrawCell;
end;

//----------------------------------------
//--- Nom   : FormShow
//--- Objet :
//----------------------------------------
procedure TFDPTableauBordDetailSolde.FormShow(Sender: TObject);
begin
 inherited;
 Self.Caption:='Solde Client : '+DonnerAuxiFromGuid (DonnerGuidFromDossier (NumDossier))+' '+DonnerLibelleDossier (NumDossier);
 If (TobSolde=nil) then InitialiserGrilleDetailSolde ();
end;

//----------------------------------------
//--- Nom   : FormClose
//--- Objet :
//----------------------------------------
procedure TFDPTableauBordDetailSolde.FormClose(Sender: TObject; var Action: TCloseAction);
begin
 inherited;
 if (TobSolde<>nil) then TobSolde.Free;
end;

//--------------------------------------------
//--- Nom   :  InitialiserGrilleDetailSolde
//--- Objet :
//--------------------------------------------
procedure TFDPTableauBordDetailSolde.InitialiserGrilleDetailSolde ();
var ChSql       : String;
    RSql        : TQuery;
    Largeur     : Integer;
begin
 //--- Création de la Tob Solde
 GrilleDetailSolde.VidePile (False);
 TobSolde:=TOB.Create('CalculSolde', Nil, -1);

 if (TobSolde<>nil) then
  begin
   ChSql:='SELECT DOS_NODOSSIER, DOS_LIBELLE from DOSSIER where DOS_CABINET="X"';
   RSql:=Opensql (ChSql, True, -1, '', True);

   while not (RSql.eof) do
    begin
     if not IsDBZero (RSql.FindField ('DOS_NODOSSIER').AsString) then
      AjouterElementGrilleSolde (TobSolde,RSql.FindField ('DOS_NODOSSIER').AsString,RSql.FindField ('DOS_LIBELLE').AsString);
     Rsql.next;
    end;

   Ferme (Rsql);
   
   if (TobSolde.Detail.Count>=2) then AjouterTotalGrilleSolde (TobSolde);
   TobSolde.PutGridDetail(GrilleDetailSolde,False,False,'DOSSIER;LIBELLE;DEBIT;CREDIT;SOLDE');
  end;

 //--- Définition de la présentation de la grille
 Largeur:=GrilleDetailSolde.Width div 6;
 GrilleDetailSolde.ColWidths[0]:=Largeur;     // Num Dossier
 GrilleDetailSolde.ColWidths[1]:=2*Largeur;   // Libelle
 GrilleDetailSolde.ColWidths[2]:=Largeur;     // Debit
 GrilleDetailSolde.ColWidths[3]:=Largeur;     // Crédit
 GrilleDetailSolde.ColWidths[4]:=Largeur;     // Solde
 HMTrad.ResizeGridColumns(GrilleDetailSolde);
end;

//----------------------------------------
//--- Nom   : AjouterElementGrilleSolde
//--- Objet :
//----------------------------------------
procedure TFDPTableauBordDetailSolde.AjouterElementGrilleSolde (Var UneTob : Tob; ChDossier,ChLibelle : String);
var ChSql,ChDebit,ChCredit,ChSolde : String;
    Rsql                           : TQuery;
    UneTobEnreg                    : Tob;
begin
 try
  ChSql:='SELECT T_TOTALDEBIT,T_TOTALCREDIT from DB'+ChDossier+'.dbo.TIERS Where T_AUXILIAIRE="'+DonnerAuxiFromGuid (DonnerGuidFromDossier (NumDossier))+'"';
  RSql:=OpenSql (ChSql,True,-1,'',True);
  if not (RSql.Eof) then
   begin
    ChDebit:=StrFMontant (RSql.FindField ('T_TOTALDEBIT').AsFloat,15,2,'',True);
    ChCredit:=StrFMontant (RSql.FindField ('T_TOTALCREDIT').AsFloat,15,2,'',True);
    ChSolde:=FormaterMontant (RSql.FindField ('T_TOTALDEBIT').AsFloat-RSql.FindField ('T_TOTALCREDIT').AsFloat);
    UneTobEnreg:=Tob.Create ('',UneTob,-1);
    UneTobEnreg.LoadFromSt ('Dossier|Libelle|Debit|Credit|Solde','|',ChDossier+'|'+ChLibelle+'|'+ChDebit+'|'+ChCredit+'|'+ChSolde,'|');
   end;
  Ferme (RSql);
 except
  PgiInfo ('Impossible d''accéder à la table Tiers.','Erreur');
 end;
end;

//---------------------------------------
//--- Nom   : AjouterTotalGrilleSolde
//--- Objet :
//---------------------------------------
procedure TFDPTableauBordDetailSolde.AjouterTotalGrilleSolde (Var UneTob : Tob);
Var TotalDebit,TotalCredit,TotalSolde : Double;
    Indice                            : Integer;
    UneTobEnreg                       : Tob;
begin
 TotalDebit:=0;TotalCredit:=0;TotalSolde:=0;

 for Indice:=0 to UneTob.Detail.Count-1 do
  begin
   TotalDebit:=TotalDebit+Valeur (UneTob.Detail[Indice].GetValue ('Debit'));
   TotalCredit:=TotalCredit+Valeur (UneTob.Detail[Indice].GetValue ('Credit'));
   TotalSolde:=Totalsolde+Valeur (UneTob.Detail[Indice].GetValue ('Solde'));
  end;

 UneTobEnreg:=Tob.create ('',UneTob,-1);
 UneTobEnreg.LoadFromSt ('Dossier|Libelle|Debit|Credit|Solde','|','Total||'+StrFMontant (TotalDebit,15,2,'',True)+'|'+StrFMontant (ToTalCredit,15,2,'',True)+'|'+FormaterMontant (ToTalSolde),'|');
end;

//----------------------------------------
//--- Nom   : GrilleDetailSoldeDblClick
//--- Objet :
//----------------------------------------
procedure TFDPTableauBordDetailSolde.GrilleDetailSoldeDblClick (Sender:TObject);
var ChCodeAuxiliaire,ChCodeGenerale : String;
    ChDossier                       : String;
begin
 inherited;

 // $$$ JP 25/04/06 - tests supplémentaires sur la ligne sélectionnée
 if (TobSolde<>nil) and (GrilleDetailSolde.Objects [0,1]<>Nil) and (GrilleDetailSolde.Row > 0) and (GrilleDetailSolde.Row <= TobSolde.Detail.Count) then
  begin
   ChDossier:=TobSolde.Detail [GrilleDetailSolde.Row-1].GetValue ('Dossier');
   if (ChDossier<>'Total') then
    begin
     ChCodeAuxiliaire:=DonnerAuxiFromGuid (DonnerGuidFromDossier (NumDossier));
     ChCodeGenerale:=DonnerGenFromAuxi (ChCodeAuxiliaire);
     UTofConsEcr.OperationsSurComptes (ChCodeGenerale,'-2','',ChCodeAuxiliaire,False,'DB'+ChDossier);
    end;
  end;
end;

//----------------------------
//--- Nom   : FetCellCanvas
//--- Objet :
//----------------------------
procedure TFDPTableauBordDetailSolde.GetCellCanvas(ACol,ARow: LongInt;  Canvas : TCanvas; AState : TGridDrawState) ;
var Texte : array[0..255] of Char;
begin
 StrPCopy(Texte,GrilleDetailSolde.Cells[0,ARow]);
 if (Texte='Total') then GrilleDetailSolde.Canvas.Font.Style:=[FsBold];
end;

//---------------------------
//--- Nom   : PostDrawCell
//--- Objet :
//---------------------------
procedure TFDPTableauBordDetailSolde.PostDrawCell(ACol,ARow: Integer; Canvas: TCanvas; AState: TGridDrawState);
begin
 if (ACol>1) and (ARow>0) then EcrireTexte (DROITE,ACol,ARow,GrilleDetailSolde)
end;

// $$$ JP 25/04/06 - FQ 11001 - F5 pour double clic
procedure TFDPTableauBordDetailSolde.GrilleDetailSoldeKeyDown (Sender:TObject; var Key:Word; Shift:TShiftState);
begin
     if Key = VK_F5 then
        GrilleDetailSoldeDblClick (nil);
end;

end.
