unit dpRepertoireTel;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Vierge, Grids, Hctrls, HSysMenu, HTB97, HPanel, UIUtil, UTob, Math,
{$IFDEF EAGLCLIENT}
  MaineAGL,
{$ELSE}
  Fe_Main,
{$ENDIF}
  HMsgBox, Menus;

type
  TFRepertoireTel = class(TFVierge)
    GdRep     :THGrid;
    PMRepert: TPopupMenu;

    procedure FormShow (Sender:TObject);
    procedure FormClose (Sender:TObject; var Action:TCloseAction);
    procedure GdRepDblClick (Sender: TObject);
    procedure FormKeyDown (Sender:TObject; var Key:Word; Shift:TShiftState);
    procedure BValiderClick (Sender:TObject);
    procedure PMRepertPopup(Sender: TObject);

  private
         GuidPer   :string;
         TobRep    :TOB;

         procedure Showfiche;
         procedure ChargeRepertoireTel;

         procedure OnTel1Click (Sender:TObject);
         procedure OnTel2Click (Sender:TObject);
         procedure OnMailClick (Sender:TObject);
  end;

procedure FicheRepertoireTel(sGuidPer : String);



implementation

uses entDP, ctiAlerte, galmenudisp, mailol;


// ouverture de la présente fiche
procedure FicheRepertoireTel (sGuidPer : String);
var
   FRepertoireTel :TFRepertoireTel;
   PP             :THPanel;
begin
     if sGuidPer='' then
     begin
          PGIInfo ('Aucun dossier sélectionné', 'Répertoire téléphonique et internet');
          exit;
     end;

     FRepertoireTel := TFRepertoireTel.Create (Application);
     FRepertoireTel.GuidPer := sGuidPer;
     PP := FindInsidePanel;
     if PP=nil then
     begin
          try
             FRepertoireTel.ShowModal;
          finally
                 FRepertoireTel.Free ;
          end ;
     end
     else
     begin
          InitInside (FRepertoireTel, PP);
          FRepertoireTel.Show;
          FRepertoireTel.SetFocus;
     end;
end;

{$R *.DFM}

procedure TFRepertoireTel.FormShow (Sender:TObject);
begin
  inherited;

  // $$$JP 23/10/03
  TobRep := nil;

  // $$$JP 16/06/05
  PMRepert.Images := FMenuDisp.SmallImages;

  ChargeRepertoireTel;
end;

procedure TFRepertoireTel.FormClose (Sender:TObject;
var
   Action   :TCloseAction);
begin
     inherited;
//     if TobRep<>Nil then
     TobRep.Free;
end;

procedure TFRepertoireTel.Showfiche;
var
   T   :TOB;
   cle :string;
begin
     // récupère la tob fille dans la col. 0
     T := TOB(GdRep.Objects[0,GdRep.Row]);

     if T.GetValue('PRINC')='>' then 
     begin
          // Fiche annuaire
          AGLLanceFiche('YY', 'ANNUAIRE', GuidPer, GuidPer, '')
     end
     else
     begin
          // Fiche contact
          // clé ANNUINTERLOC : ANI_GUIDPER;ANI_NOCONT
         (* MCD 12/2005cle := GuidPer + ';' + IntToStr (T.GetValue ('ANI_NOCONT'));
          AGLLanceFiche ('YY', 'ANNUINTERLOC', cle, cle, GuidPer+';MONOFICHE=TRUE');*)
          cle := T.GetValue ('C_TYPECONTACT') + ';'+ T.GetValue ('C_AUXILIAIRE') + ';'+ IntToStr (T.GetValue ('C_NUMEROCONTACT'));
          AGLLanceFiche ('YY', 'YYCONTACT', cle, cle, ';GUIDPER='+GuidPer+';MONOFICHE=TRUE');
     end;

     // Actualisation
     ChargeRepertoireTel;
end;

procedure TFRepertoireTel.GdRepDblClick (Sender:TObject);
begin
     inherited;

     ShowFiche;
end;

procedure TFRepertoireTel.ChargeRepertoireTel;
var
   larg  :integer;
begin
     // Tob répertoire (fiche principale de la soc + contacts)
     if TobRep <> nil then
         TobRep.ClearDetail // $$$ JP 23/10/2003: détruire au préalable la TOB déjà chargée
     else
         TobRep := TOB.Create ('', Nil, -1);

    (* mcd 12/2005  TobRep.LoadDetailFromSQL (
             'SELECT 0 AS ANI_NOCONT, ">" AS PRINC, ANN_CV AS ANI_CV, ANN_NOM1 AS ANI_NOM, ANN_NOM2 AS ANI_PRENOM,'
            +' ANN_TEL1 AS ANI_TEL1, ANN_TEL2 AS ANI_TEL2, ANN_FAX AS ANI_FAX, ANN_EMAIL AS ANI_EMAIL'
            +' FROM ANNUAIRE WHERE ANN_GUIDPER="'+GuidPer+'"'
            +' UNION '
            +'SELECT ANI_NOCONT, "" AS PRINC, ANI_CV, ANI_NOM, ANI_PRENOM, ANI_TEL1, ANI_TEL2, ANI_FAX, ANI_EMAIL'
            +' FROM ANNUINTERLOC'
            +' WHERE ANI_GUIDPER="'+GuidPer+'"'
            +'ORDER BY ANI_NOM, ANI_PRENOM'); *)
     TobRep.LoadDetailFromSQL (
             'SELECT 0 AS C_NUMEROCONTACT, ">" AS PRINC, "" AS C_TYPECONTACT,"" AS C_AUXILIAIRE, ANN_CV AS C_CIVILITE, ANN_NOM1 AS C_NOM, ANN_NOM2 AS C_PRENOM,'
            +' ANN_TEL1 AS C_TELEPHONE, ANN_TEL2 AS C_TELEX, ANN_FAX AS C_FAX, ANN_EMAIL AS C_RVA'
            +' FROM ANNUAIRE WHERE ANN_GUIDPER="'+GuidPer
            +'" UNION '
            +'SELECT C_NUMEROCONTACT, "" AS PRINC, C_TYPECONTACT,C_AUXILIAIRE,C_CIVILITE, C_NOM, C_PRENOM, C_TELEPHONE, C_TELEX, C_FAX, C_RVA'
            +' FROM CONTACT'
            +' WHERE C_GUIDPER="'+GuidPer
            +'" ORDER BY C_NOM, C_PRENOM');

     // affichage
    //mcd 12/2005 TobRep.PutGridDetail(GdRep, False, False, 'PRINC;ANI_CV;ANI_NOM;ANI_PRENOM;ANI_TEL1;ANI_TEL2;ANI_FAX;ANI_EMAIL');
     TobRep.PutGridDetail(GdRep, False, False, 'PRINC;C_CIVILITE;C_NOM;C_PRENOM;C_TELEPHONE;C_TELEX;C_FAX;C_RVA');
     GdRep.RowCount := Max(TobRep.Detail.Count+1,2); // ou TobRep.FillesCount(0)+1;

     larg := GdRep.Width div 40;
     GdRep.ColWidths[0] := larg;   // Princ
     GdRep.ColWidths[1] := 4*larg; // Civ.
     GdRep.ColWidths[2] := 8*larg; // Nom
     GdRep.ColWidths[3] := 6*larg; // Prénom
     GdRep.ColWidths[4] := 4*larg; // Téléphone
     GdRep.ColWidths[5] := 4*larg; // Autre tél.
     GdRep.ColWidths[6] := 4*larg; // Télécopie
     GdRep.ColWidths[7] := 6*larg; // Email

     HMTrad.ResizeGridColumns (GdRep);
end;

procedure TFRepertoireTel.FormKeyDown (Sender:TObject; var Key:Word; Shift:TShiftState);
begin
     inherited;

     if Key = VK_F5 then
        ShowFiche;
//        GdRepDblClick (Sender);
end;

procedure TFRepertoireTel.BValiderClick (Sender:TObject);
begin
     inherited;

     ShowFiche;
end;

procedure TFRepertoireTel.PMRepertPopup (Sender:TObject);
var
   T        :TOB;
   MItem    :TMenuItem;
   strTel   :string;
   strMail  :string;
begin
     PMRepert.Items.Clear;

     // On appel le téléphone demandé
     T := TOB (GdRep.Objects [0,GdRep.Row]);
     if T <> nil then
     begin
          // Menu contextuel pour appeller les téléphones (si CTI actif)
          if VH_DP.ctiAlerte <> nil then
          begin
               strTel := Trim (T.GetString ('C_TELEPHONE'));
               if strTel <> '' then
               begin
                    MItem := TMenuItem.Create (PMRepert);
                    if MItem <> nil then
                    begin
                         MItem.Caption    := 'Appeler le ' + strTel;
                         MItem.OnClick    := OnTel1Click;
                         MItem.ImageIndex := cImgMenuDispTel;
                         PMRepert.Items.Add (MItem);
                    end;
               end;

               strTel := Trim (T.GetString ('C_TELEX'));
               if strTel <> '' then
               begin
                    MItem := TMenuItem.Create (PMRepert);
                    if MItem <> nil then
                    begin
                         MItem.Caption    := 'Appeler le ' + strTel;
                         MItem.OnClick    := OnTel2Click;
                         MItem.ImageIndex := cImgMenuDispTel;
                         PMRepert.Items.Add (MItem);
                    end;
               end;
          end;

          // Envoie d'email
          strMail := Trim (T.GetString ('C_RVA'));
          if strMail <> '' then
          begin
               MItem := TMenuItem.Create (PMRepert);
               if MItem <> nil then
               begin
                    MItem.Caption := 'Ecrire à ' + strMail;
                    MItem.OnClick := OnMailClick;
                    MItem.ImageIndex := cImgMenuDispMail;
                    PMRepert.Items.Add (MItem);
               end;
          end;
     end;
end;

procedure TFRepertoireTel.OnTel1Click (Sender:TObject);
var
   T        :TOB;
   strTel   :string;
begin
     T := TOB (GdRep.Objects [0,GdRep.Row]);
     if T <> nil then
     begin
          strTel := Trim (T.GetString ('C_TELEPHONE'));
          if PgiAsk ('Appeler le ' + strTel + ' ?') = mrYes then
             VH_DP.ctiAlerte.MakeCall (strTel);
     end;
end;

procedure TFRepertoireTel.OnTel2Click (Sender:TObject);
var
   T        :TOB;
   strTel   :string;
begin
     T := TOB (GdRep.Objects [0,GdRep.Row]);
     if T <> nil then
     begin
          strTel := Trim (T.GetString ('C_TELEX'));
          if PgiAsk ('Appeler le ' + strTel + ' ?') = mrYes then
             VH_DP.ctiAlerte.MakeCall (strTel);
     end;
end;

procedure TFRepertoireTel.OnMailClick (Sender:TObject);
var
   T        :TOB;
   strMail  :string;
begin
     T := TOB (GdRep.Objects [0,GdRep.Row]);
     if T <> nil then
     begin
          strMail := Trim (T.GetString ('C_RVA'));
          SendMail ('', strMail, '', nil, '', FALSE);
     end;
end;

end.
