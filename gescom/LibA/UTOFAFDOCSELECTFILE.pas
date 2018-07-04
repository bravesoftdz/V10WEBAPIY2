{***********UNITE*************************************************
Auteur  ...... : JP
Créé le ...... : 26/07/2002
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : AFDOCSELECTFILE ()
Mots clefs ... : TOF;AFDOCSELECTFILE
*****************************************************************}
Unit UTOFAFDOCSELECTFILE ;

Interface

Uses StdCtrls, 
     Controls, 
     Classes, 
{$IFNDEF EAGLCLIENT}
     db, 
     dbtables, 
{$ENDIF}
     forms, 
     sysutils, 
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     UTOF ;

Type
  TOF_AFDOCSELECTFILE = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnClose                  ; override ;

    procedure OnDoubleClicList (Sender:TObject);
  private
         FileList      :TListBox;
         strOrigine    :string;
         strType       :string;
         bSelected     :boolean;
  end ;

function AFLanceFiche_DocSelectFile (Range, Argument:string):string;

Implementation

uses
{$IFNDEF EAGLCLIENT}
         fe_main, AffaireOle,
{$ELSE}
       MaineAgl,
{$ENDIF}
    vierge, pgienv, utilGa;


procedure TOF_AFDOCSELECTFILE.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_AFDOCSELECTFILE.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_AFDOCSELECTFILE.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_AFDOCSELECTFILE.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_AFDOCSELECTFILE.OnArgument (S:string);
var
   SearchRec                    :TSearchRec;
   Critere                      :string;
// $$$ JP 07/07/03   AppliPathStd, AppliPathDat   :string;
   strPathSearch                :string;
begin
     inherited;

     // $$$ JP 07/07/03: amélioration identification STD ET DAT
//{$IFNDEF EAGLCLIENT}
  //       if V_PGI_ENV <> nil then
    //     begin
//{$IFDEF GIGI}
  //            AppliPathStd := V_PGI_ENV.PathStd + '\GIS5';
    //          AppliPathDat := V_PGI_ENV.PathDat + '\GIS5';
//{$ELSE}
  //            AppliPathStd := V_PGI_ENV.PathStd + '\GAS5';
    //          AppliPathDat := V_PGI_ENV.PathDat + '\GAS5';
//{$ENDIF}
//         end
  //       else
    //     begin
//              // V_PGI_ENV inexistant, pas normal!
  //            AppliPathStd := 'C:\PGI00\STD';
    //          AppliPathDat := 'C:\PGI01\STD';
       //  end;
//{$ENDIF}

     // Listbox des fichiers à énumérer
     FileList := TListBox (GetControl ('LBFICHIER'));
     FileList.OnDblClick := OnDoubleClicList;
     bSelected := FALSE;

     // Arguments: répertoire à énumérer
     Critere := (Trim (ReadTokenSt (S)));
     while (Critere <>'') do
     begin
          // Origine: std (cegid), dat (client) ou dossier (client mais uniquement pour cette base)
          if (copy (Critere,1,7) = 'ORIGINE') then
             strOrigine := Copy (Critere, 9, Length (Critere)-8);

          // Type de fichier: word, excel ...
          if (copy (Critere,1,4) = 'TYPE') then
             strType := Copy (Critere, 6, Length (Critere)-5);

          // Paramètre suivant
          Critere := (Trim (ReadTokenSt (S)));
     end;

     // Remplit la liste avec les fichiers existants, plus le premier élément pour nouveau fichier
     FileList.Items.Add ('NOUVEAU DOCUMENT');
     if strOrigine = '$STD' then
         strPathSearch := GetAppliPathStd + '\*.' // $$$ JP 07/07/03: amélioration identification STD ET DAT
     else
//         FindFirst (doc', faAnyFile, SearchRec)
  //   else
         if strOrigine = '$DAT' then
             strPathSearch := GetAppliPathDat + '\*.' // $$$ JP 07/07/03: amélioration identification STD ET DAT
//             FindFirst (doc', faAnyFile, SearchRec)
         else
             if strOrigine = V_PGI_ENV.NoDossier then
             strPathSearch := V_PGI_ENV.PathDos + '\*.';
//                FindFirst (doc', faAnyFile, SearchRec);
     if strType = 'WOR' then
         strPathSearch := strPathSearch + 'doc'
     else
         if strType = 'EXC' then
             strPathSearch := strPathSearch + 'xls'
         else
             strPathSearch := strPathSearch + 'txt';

     FindFirst (strPathSearch, faAnyFile, SearchRec);
     repeat
           if SearchRec.Name <> '' then
              FileList.Items.Add (SearchRec.Name);
     until FindNext (SearchRec) <> 0;
     sysutils.FindClose (SearchRec);
end;

procedure TOF_AFDOCSELECTFILE.OnClose;
begin
     if bSelected = FALSE then
         TFVierge (Ecran).Retour := ''
     else
         if FileList.ItemIndex > 0 then
             TFVierge (Ecran).Retour := FileList.Items.Strings [FileList.ItemIndex]
         else
             TFVierge (Ecran).Retour := '$NEW';

     inherited;
end;

procedure TOF_AFDOCSELECTFILE.OnDoubleClicList (Sender:TObject);
begin
     bSelected := TRUE;
     Ecran.Close;
end;

function AFLanceFiche_DocSelectFile (Range, Argument:string):string;
begin
     Result := AGLLanceFiche ('AFF','AFDOCSELECTFILE', Range, '', Argument) ;
end;



Initialization
              registerclasses ( [ TOF_AFDOCSELECTFILE ] ) ;
end.

