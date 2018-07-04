{***********UNITE*************************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 03/10/2006
Modifié le ... :   /  /    
Description .. :
Mots clefs ... :
*****************************************************************}
unit UFMaquette;
/////////////////////////////////////////////////////////////////
interface
/////////////////////////////////////////////////////////////////
uses
   Windows, Hent1, hmsgbox, UYFileSTD, ugedfiles, HCtrls,
	Forms, Classes, Controls, OleCtnrs, SysUtils, UTOB, AGLUtilOLE,
   comObj, extctrls, TlHelp32;
/////////////////////////////////////////////////////////////////
type
	TPActionFiche = ^TActionFiche;
   TPForm = ^TForm;

type TDocChange = procedure(bDocChange_p : boolean) of Object;
/////////////////////////////////////////////////////////////////

type
   TFMaquette = class(TForm)

    		PContainer_c: TPanel;

    		procedure FormClose(Sender: TObject; var Action: TCloseAction);
    		procedure FormDestroy(Sender: TObject);

      protected

      public

      	PDocChange_c : TDocChange;

         property  OnDocChange_c : TDocChange read PDocChange_c  write PDocChange_c;

         function  InitContainer : boolean;
         procedure DocAttache;
			function  DocIsModified : boolean;

      private

    		OleContainer_c : TOleContainer;
         bDocChange_c, bWithAskEnreg_c, bNewWord_c : boolean;
         MyAppli_c, MyDoc_c : oleVariant;
         sModele_c, sDocPath_c : string;
end;


/////////////////////////////////////////////////////////////////
function LanceLaMaquette(sDocPath_p, sDocTitre_p : string;
								 bWithAskEnreg_p, bWithModele_p : boolean;
                         PDocChange_p : TDocChange) : TModalResult;

function LanceLeDoc(sDocPath_p : string; bLectureSeule_p : boolean = false) : boolean;
/////////////////////////////////////////////////////////////////

implementation

/////////////////////////////////////////////////////////////////
{$R *.DFM}
/////////////////////////////////////////////////////////////////


{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 29/08/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
function LanceLeDoc(sDocPath_p : string; bLectureSeule_p : boolean = false) : boolean;
begin
   result := false;
   if sDocPath_p = '' then
   begin
      PGIInfo('Le nom du document n''est pas renseigné', 'Lancement de Word');
   end
   else if FileExists(sDocPath_p) then
   begin
      OpenDoc(sDocPath_p, wdWindowStateMaximize, wdPrintView, bLectureSeule_p);
      result := true;
   end
   else
      PGIInfo('Le document n''existe pas :'#10#13 + sDocPath_p, 'Lancement de Word');
end;

/////////////////////////////////////////////////////////////////
{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 20/07/2007
Modifié le ... :   /  /    
Description .. :
Mots clefs ... : 
*****************************************************************}
function LanceLaMaquette(sDocPath_p, sDocTitre_p : string;
								 bWithAskEnreg_p, bWithModele_p : boolean;
								 PDocChange_p : TDocChange) : TModalResult;
var
   FMaquette_l : TFMaquette;
begin
   result := mrNone;
   if not FileExists(sDocPath_p) then exit;

   FMaquette_l := TFMaquette.Create(Application);

   FMaquette_l.Caption := sDocTitre_p;
   FMaquette_l.sDocPath_c := sDocPath_p;
   FMaquette_l.bWithAskEnreg_c := bWithAskEnreg_p;

   if bWithModele_p then
      FMaquette_l.sModele_c := GetWordTemplateDirectory + '\CEGIDJUR1.DOT';

   if FMaquette_l.InitContainer then
      FMaquette_l.DocAttache;

   FMaquette_l.Visible := false;
   FMaquette_l.PDocChange_c := PDocChange_p;

   try
      FMaquette_l.ShowModal;
   finally
      result := FMaquette_l.ModalResult;
//      FMaquette_l.Free;
   end;
end;


/////////////////////////////////////////////////////////////////

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 20/07/2007
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
function TFMaquette.InitContainer : boolean;
var
   bInitOK_l : boolean;
begin
   bInitOK_l := true;
   OleContainer_c := TOlecontainer.Create(PContainer_c);
   OleContainer_c.AutoActivate := aaManual;
   OleContainer_c.Parent := PContainer_c;
   OleContainer_c.Align := alClient;
   OleContainer_c.Visible := true;

   try
      // y a t'il une instance de word existante?
	   MyAppli_c := GetActiveOleObject('Word.Application');
      bNewWord_c := false;
   except
      try
         // création d'une nouvelle instance de word      
         MyAppli_c := CreateOleObject('Word.Application');
         bNewWord_c := true;
      except
         bInitOK_l := false;
      end;
   end;

   result := bInitOK_l;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 02/12/2006
Modifié le ... :   /  /    
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TFMaquette.DocAttache;
begin
   OleContainer_c.CreateObject('Word.Document', false);
   OleContainer_c.Run; //active l'appli serveur sans activer l'objet ==> + rapide pour le reste.
   OleContainer_c.Realign;

	if (sDocPath_c = '') or not FileExists(sDocPath_c) then
   begin
      OleContainer_c.CreateObject('Word.Document', false);
   end
   else
   begin
	   OleContainer_c.CreateObjectFromFile(sDocPath_c, false);
   end;

   OleContainer_c.DoVerb(ovShow);
	OleContainer_c.Realign;

   MyAppli_c := OleContainer_c.OleObject.Application;
   MyDoc_c := MyAppli_c.ActiveDocument;
   MyDoc_c.Activate;
   MyDoc_c.ActiveWindow.View.Type := wdPrintView;
	OleContainer_c.Realign;

//   MyDoc_c.AttachedTemplate := sModele_c;
   bDocChange_c := false;

end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 02/12/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
function  TFMaquette.DocIsModified : boolean;
begin
	bDocChange_c := (not MyDoc_c.Saved);
	result := bDocChange_c;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 03/12/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TFMaquette.FormClose(Sender: TObject; var Action: TCloseAction);
begin
	inherited;
   if @PDocChange_c = nil then
   begin
   	PGIError('ERREUR : procédure "TFMaquette.PDocChange_c" non affectée');
   end
   else
   begin
	   MyDoc_c.Activate;

   	if DocIsModified then
      begin
		   if bWithAskEnreg_c then
		   begin
   			if (PGIAsk('Le document a été modifié. Voulez-vous l''enregistrer?') = mrYes) then
            begin
               OleContainer_c.SaveAsDocument(sDocPath_c);
               PDocChange_c(true);
            end;
         end
         else
         begin
            OleContainer_c.SaveAsDocument(sDocPath_c);
            PDocChange_c(true);
         end;
      end;
   end;

   OleContainer_c.Close();
   OleContainer_c.DestroyObject;   
   if bNewWord_c then
	   MyAppli_c.Quit;
//   Action := caFree;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 20/07/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TFMaquette.FormDestroy(Sender: TObject);
begin
	FreeAndNil(OleContainer_c);
end;



/////////////////////////////////////////////////////////////////



end.
