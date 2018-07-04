{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 23/11/2005
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : PGTESTFORMATION ()
Mots clefs ... : TOF;PGTESTFORMATION
*****************************************************************
PT1 28/09/2007  V_7  FL  Adaptation cursus + accès assistant
PT2 17/10/2007  V_7  FL  Correction bugs mémoire + raccourcis F9,F10
PT3 18/10/2007  V_7  FL  Mise en forme de la grille + Réduction sensible du nombre de requêtes
PT4 24/10/2007  V_8  FL  Gestion des cursus favoris + Correction des exports
PT5 04/12/2007  V_8  FL  Correction du cumul des honoraires + Modification de l'affichage des sous-niveaux
}
Unit UTofPGEM_Previsionnel ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
     Graphics,
     HTB97,
{$IFNDEF EAGLCLIENT}
     db,
     dbtables,
     mul,
     FE_Main,
     Fiche,
     EdtREtat,
{$else}
     eMul,
     MainEAGL,
     UTilEAGL,
{$ENDIF}
     forms,
     uTob,
     sysutils,
     ComCtrls,
     ParamSoc,
     HCtrls,
     HEnt1,
     Vierge,
     HMsgBox,
     UTobDebug,
     HSysMenu,
     Grids,
     ed_tools,
     P5Util,
     ExtCtrls,
     Menus,
     ShellAPI,
     HPanel,
     windows,
     EntPaie,
     MailOl,
     PGoutilsFormation,
     lookUp,
     Dialogs,
     UTOF ;

Type
  TOF_PGTESTFORMATION = Class (TOF)
    procedure OnArgument (S : String ) ; override ;
    private
    Grille : THGrid;
    Millesime : String;
    DateSortieSal : TDateTime;
    TypeUtilisat,Utilisateur : String;
    BCherche : TToolBarButton97;
    Consultation : Boolean;
    procedure GrilleGetCellCanvas( ACol, ARow : LongInt; Canvas : TCanvas; AState : TGridDrawState) ;
    procedure AccesFormation(Sender : TObject);
    procedure AccesSalarie(Sender : TObject);
    procedure AccesCursus(Sender : TObject); //PT1
    procedure AffichageParSalarie;
    procedure AffichageParFormation;
    procedure AffichageParCursus; //PT1
    //procedure AffichageParRespFormation(Sender : Tobject);		//PT5
    procedure AffichageParLibEmploi;
    procedure ChangeRegroupement(Sender : TObject);
    procedure GrilleDblClick(Sender : TObject);
    procedure BChercheClick(Sender : TObject);
    procedure ChangeCritere(Sender : Tobject);
    procedure SalarieElipsisClick(Sender: TObject);
    procedure RespElipsisClick(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure ExportGrille(Sender : Tobject);
  end ;

Implementation

procedure TOF_PGTESTFORMATION.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  case Key of
    VK_F9,VK_F10 : If BCherche.Enabled = True then BChercheClick(Nil); //PT2
      end;
end;


procedure TOF_PGTESTFORMATION.OnArgument (S : String ) ;
var Combo : THValComboBox;
    Bt : TToolBarButton97;
//    Q : TQuery;
    Edit : THEdit;
    Check : TCheckBox;
begin
  Inherited ;
  //PT1 - Début
  If ExisteSQL('SELECT PSE_SALARIE FROM DEPORTSAL WHERE PSE_RESPONSFOR="'+V_PGI.UserSalarie+'"') then TypeUtilisat := 'R'
  else TypeUtilisat := 'S';
  //PT1 - Fin
  TFVierge(Ecran).OnKeyDown := FormKeyDown;
  Edit := THEdit(GetControl('PSA_SALARIE'));
  if Edit <> nil then Edit.OnElipsisClick := SalarieElipsisClick;
  DateSortieSal := V_PGI.DateEntree;
  Millesime := RendMillesimeEManager;
  Utilisateur := V_PGI.UserSalarie;
  If Millesime = '0000' then
  begin
     Consultation := True;
     //Q := OpenSQL('SELECT PFE_MILLESIME FROM EXERFORMATION ORDER BY PFE_MILLESIME DESC',True);
     //If Not Q.Eof then Millesime := Q.FindField('PFE_MILLESIME').AsString;
     //Ferme(Q);
     Millesime := RendMillesimeEManager; //PT1
  end
  else Consultation := False;
  Grille := THGrid(GetControl('GRILLE'));
  Grille.ColAligns[2] := taRightJustify;
  Grille.ColAligns[3] := taRightJustify;
  Grille.ColAligns[4] := taRightJustify;
  Grille.ColFormats[2] := '# ##0';
  Grille.ColFormats[3] := '# ##0.00';
  Grille.ColFormats[4] := '# ##0';
  Grille.ColFormats[5] := '';
  Grille.ColFormats[6] := '';
  Grille.GetCellCanvas := GrilleGetCellCanvas;
  Grille.OnDblClick := GrilleDblClick;
  Bt := TToolBarButton97(GetControl('BINSERT'));
  If Consultation then Bt.Visible := False
  else
  begin
    If S = 'FOR' then
    Begin
         Bt.OnClick := AccesFormation;
         //PT1 - Début
         SetControlVisible('FAVORIS',True);
         Check := TCheckBox(GetControl('FAVORIS'));
         If Check <> Nil then Check.OnClick := ChangeCritere;
         //PT1 - Fin
         If ExisteSQL('SELECT PCC_CODESTAGE FROM CURSUSSTAGE WHERE PCC_MILLESIME="'+Millesime+'" AND PCC_CURSUS="---" AND PCC_LIBELLE="'+V_PGI.UserSalarie+'"') then
               SetControlChecked('FAVORIS',True)
         Else
               SetControlChecked('FAVORIS',False);
    End;
    If S = 'SAL' then Bt.OnClick := AccesSalarie;
    If S = 'CUR' then
    Begin
          Bt.OnClick := AccesCursus; //PT1
         //PT4 - Début
         SetControlVisible('FAVORIS',True);
         Check := TCheckBox(GetControl('FAVORIS'));
         If Check <> Nil then Check.OnClick := ChangeCritere;

         If ExisteSQL('SELECT PCC_CODESTAGE FROM CURSUSSTAGE WHERE PCC_MILLESIME="'+Millesime+'" AND PCC_CURSUS="-C-" AND PCC_LIBELLE="'+V_PGI.UserSalarie+'"') then
               SetControlChecked('FAVORIS',True)
         Else
               SetControlChecked('FAVORIS',False);
         //PT4 - Fin
    End;
  end;
  BCherche := TToolBarButton97(GetControl('BCHERCHE'));
  If BCherche <> Nil then BCherche.OnClick := BChercheClick;
  Combo := THValComboBox(GetControl('CREGROUPEMENT'));
  If Combo <> Nil then
  begin
   Combo.OnChange := ChangeRegroupement;
   If S = 'SAL' then
   begin
        Combo.Items.Delete(3); //CUR //PT1
        Combo.Items.Delete(2); //FOR
   end
   else
   begin
        Combo.Enabled := False;
   end;
  end;

  Edit := THEdit(GetControl('PSA_SALARIE'));
  If Edit <> Nil then Edit.OnChange := ChangeCritere;
  Combo := THValComboBox(GetControl('PSA_LIBELLEEMPLOI'));
  If Combo <> Nil then Combo.OnChange := ChangeCritere;
  //Combo := THValComboBox(GetControl('PSA_ETABLISSEMENT')); //PT1 Plus d'établissement
  //If Combo <> Nil then Combo.OnChange := ChangeCritere;
  Edit := THEdit(GetControl('RESPONSFOR'));
  If Edit <> Nil then
  begin
    Edit.OnChange := ChangeCritere;
    Edit.OnElipsisClick := RespElipsisClick;
  end;
  SetControlCaption('LRESPONS','');
  SetControlText('CREGROUPEMENT',S);
  If S = 'SAL' then TFVierge(Ecran).Caption := 'Inscriptions par salarié pour le prévisionnel '+Millesime
  else If S = 'FOR' Then TFVierge(Ecran).Caption := 'Inscriptions par formation pour le prévisionnel '+Millesime //PT1
  else If S = 'CUR' Then TFVierge(Ecran).Caption := 'Inscriptions par cursus pour le prévisionnel '+Millesime; //PT1
  UpdateCaption(Ecran);
  SetControlProperty('PSA_LIBELLEEMPLOI','Plus', ' AND CC_CODE IN (SELECT PSA_LIBELLEEMPLOI FROM SALARIES LEFT JOIN DEPORTSAL ON PSA_SALARIE=PSE_SALARIE '+
  'WHERE (PSA_DATESORTIE IS NULL OR PSA_DATESORTIE>="'+UsdateTime(V_PGI.DateEntree)+'" OR PSA_DATESORTIE<="'+UsDateTime(IDate1900)+'") AND '+
  //PT1 - Début
  //'(PSE_RESPONSFOR="'+V_PGI.UserSalarie+'" OR PSE_RESPONSFOR IN (SELECT PGS_RESPONSFOR FROM SERVICES LEFT JOIN SERVICEORDRE ON PGS_CODESERVICE=PSO_CODESERVICE'+
  //' WHERE PSO_SERVICESUP IN (SELECT PGS_CODESERVICE FROM SERVICES WHERE PGS_RESPONSFOR="'+V_PGI.UserSalarie+'"))))');
  AdaptByRespEmanager (TypeUtilisat,'PSE',Utilisateur,False)+')');
  //PT1 -Fin
  Check := TCheckBox(GetControl('MULTINIVEAU'));
  //PT5 - Début
  //If Check <> Nil then Check.OnClick := AffichageParRespFormation;
  If Check <> Nil Then Check.OnClick := ChangeCritere;
  //PT5 - Fin
  Bt := TToolBarButton97(GetControl('BEXPORT'));
  If Bt <> Nil then
  begin
       Bt.Visible := True;
       Bt.OnClick := ExportGrille;
       Bt.Hint := 'Exporter';
       Bt.ShowHint := True;
  end;

  //PT4 - Début
  If TypeUtilisat = 'R' Then
  Begin
       // Responsable actuel par défaut
       SetControlText('RESPONSFOR', V_PGI.UserSalarie);

       // Bloque la sélection des responsables s'il n'y en a pas en dessous
       If Not ExisteSQL('SELECT PGS_RESPONSFOR FROM SERVICES WHERE PGS_CODESERVICE IN '+
           '(SELECT PSO_CODESERVICE FROM SERVICES LEFT JOIN SERVICEORDRE ON PGS_CODESERVICE=PSO_SERVICESUP AND'+
           ' PGS_RESPONSFOR="'+V_PGI.UserSalarie+'")') Then
       Begin
          SetControlEnabled('RESPONSFOR', False);
          SetControlEnabled('MULTINIVEAU', False);
       End;
       SetControlEnabled('BCHERCHE', False);
  End;
  //PT4 - Fin
end ;

procedure TOF_PGTESTFORMATION.ChangeRegroupement(Sender : TObject);
begin
     (*If GetCheckBoxState('MULTINIVEAU') = CbChecked then //PT5
     begin
          AffichageParRespFormation(Nil);
          Exit
     end;*)
     If GetControlText('CREGROUPEMENT') = 'FOR' then AffichageParFormation
     else if GetControlText('CREGROUPEMENT') = 'LIBEMP' then Affichageparlibemploi
     else if GetControlText('CREGROUPEMENT') = 'SAL' Then AffichageParSalarie //PT1
     Else if GetControlText('CREGROUPEMENT') = 'CUR' Then AffichageParCursus; //PT1

end;

procedure TOF_PGTESTFORMATION.GrilleGetCellCanvas( ACol, ARow : LongInt; Canvas : TCanvas; AState : TGridDrawState) ;
var Col : Integer;
begin
	(* //PT5
     If GetCheckBoxState('MULTINIVEAU') = CbChecked then
     begin
           If Grille.CellValues[9,ARow] = 'R' then
           begin
                If Grille.Row = ARow then Canvas.Font.Color := ClBlue;
                Canvas.Font.Style := [fsBold];
                Canvas.Brush.Color := ClBtnFace;
           end;
     end
     ELSE*) 
     If GetCheckBoxState('MULTINIVEAU') = CbChecked then
     	Col := 8
     Else 
     	Col := 7;
     
     
     If GetControlText('CREGROUPEMENT') = 'FOR' then
     begin
           If (Grille.CellValues[Col,ARow] = 'F') and Not(Grille.IsSelected(ARow)) then
           begin
                If Grille.Row = ARow then Canvas.Font.Color := ClBlue;
                Canvas.Font.Style := [fsBold];
                Canvas.Brush.Color := Cl3DLight;
                Canvas.Brush.Color := ClBtnFace;
           end;
     end
     //PT1 - Début
     ELSE If GetControlText('CREGROUPEMENT') = 'CUR' then
     begin
           If (Grille.CellValues[Col,ARow] = 'C') and Not(Grille.IsSelected(ARow)) then
           begin
                If Grille.Row = ARow then Canvas.Font.Color := ClBlue;
                Canvas.Font.Style := [fsBold];
                Canvas.Brush.Color := Cl3DLight;
                Canvas.Brush.Color := ClBtnFace;
           end;
     end
     //PT1 - Fin
     else
      If (GetControlText('CREGROUPEMENT') = 'RESPSAL') or (GetControlText('CREGROUPEMENT') = 'RESPFOR') then
     begin
           If Grille.CellValues[Col,ARow] = 'R' then
           begin
                If Grille.Row = ARow then Canvas.Font.Color := ClBlue;
                Canvas.Font.Style := [fsBold];
                Canvas.Brush.Color := ClBtnFace;
           end;
     end
     else
      If GetControlText('CREGROUPEMENT') = 'LIBEMP' then
     begin
           If Grille.CellValues[Col,ARow] = 'L' then
           begin
                If Grille.Row = ARow then Canvas.Font.Color := ClBlue;
                Canvas.Font.Style := [fsBold];
                Canvas.Brush.Color := ClBtnFace;
           end;
     end
     else
     begin
          If (Grille.CellValues[Col,ARow] = 'S') and Not(Grille.IsSelected(ARow)) then
          begin
               If Grille.Row = ARow then Canvas.Font.Color := ClBlue;
               Canvas.Font.Style := [fsBold];
               Canvas.Brush.Color := ClBtnFace;
          end;
     end;
end;

procedure TOF_PGTESTFORMATION.AccesFormation(Sender : TObject);
begin
     AglLanceFiche('PAY', 'EM_MULSTAGE', '', '', ''); // inscriptions a partir des formations
     ChangeRegroupement(Nil);
end;

procedure TOF_PGTESTFORMATION.AccesSalarie(Sender : TObject);
begin
     AglLanceFiche('PAY', 'EM_MULSALARIE', '', '', '');
     ChangeRegroupement(Nil);
end;

//PT1 - Début
procedure TOF_PGTESTFORMATION.AccesCursus(Sender: TObject);
begin
     AglLanceFiche('PAY', 'EM_MULCATCURSUS', '', '', 'INSC');
     ChangeRegroupement(Nil);
end;
//PT1 - Fin

procedure TOF_PGTESTFORMATION.AffichageParSalarie;
var Q : TQuery;
    TG,TGrille,TFor,TForN : Tob;
    i : Integer;
    Salarie,Where : String;
    NbHeures,Cout : Double;
    NbInsc : Integer;
    TSomme,TS : Tob;
    TotHeures,TotSal,TotInsc,Tothono : Double;
    NbCol, Col : Integer; //PT5
begin
  TotSal    := 0;
  TotInsc   := 0;
  TotHeures := 0;
  TotHono   := 0;
  NbInsc    := 0;
  Cout      := 0;
  NbHeures  := 0;

  For i := 0 to Grille.RowCount - 1 do Grille.Rows[i].Clear;
  Grille.RowCount := 3;
  //PT5 - Début
  If (GetCheckBoxState('MULTINIVEAU')=CbChecked) Then
    NbCol := 12
  Else
    NbCol := 11; //PT3
  Grille.ColCount := NbCol;
  //PT5 - Fin

  //PT3 - Début
  // Cas particulier des inscriptions non nominatives
  // ------------------------------------------------
  If GetControlText('PSA_SALARIE') = '' Then
  Begin
       Where := '';
       If GetControlText('PSA_LIBELLEEMPLOI') <> '' Then Where := Where + ' AND PFI_LIBEMPLOIFOR="'+GetControlText('PSA_LIBELLEEMPLOI')+'"';
       Where := Where + ' AND ' + AdaptByRespEmanager (TypeUtilisat,'PFI',Utilisateur,(GetCheckBoxState('MULTINIVEAU')=CbChecked)); //PT5

       Q := OpenSQL('SELECT "NULL" AS PSA_SALARIE,"" AS PSA_LIBELLE,"" AS PSA_PRENOM,"" AS PSA_LIBELLEEMPLOI,'+
                    'PFI_SALARIE,PFI_LIBELLE,PFI_MOTIFINSCFOR,PFI_NIVPRIORITE,PFI_CODESTAGE,PFI_AUTRECOUT,PFI_NATUREFORM,PFI_RESPONSFOR,PFI_NBINSC,PFI_DUREESTAGE,PFI_ETATINSCFOR,PFI_MOTIFETATINSC '+
                    'FROM INSCFORMATION '+
                    'WHERE PFI_MILLESIME="'+Millesime+'" AND PFI_CODESTAGE<>"--CURSUS--" AND PFI_SALARIE=""'+Where+
                    ' ORDER BY PFI_LIBELLE',True);

       TForN := Tob.Create('Table',Nil,-1);
       TForN.LoadDetailDB('Table','','',Q,False);
       Ferme(Q);
  end;
  //PT3 - Fin

  // Autres inscriptions
  // -------------------
  Where := '';
  If GetControlText('PSA_LIBELLEEMPLOI') <> '' Then Where := Where + ' AND PSA_LIBELLEEMPLOI="'+GetControlText('PSA_LIBELLEEMPLOI')+'"';
  If GetControlText('PSA_SALARIE') <> '' Then Where := Where + ' AND PSE_SALARIE="' + GetControlText('PSA_SALARIE')+'"';
  Where := Where + ' AND ' + AdaptByRespEmanager (TypeUtilisat,'PSE',Utilisateur,(GetCheckBoxState('MULTINIVEAU')=CbChecked)); //PT5

  Q := OpenSQL('SELECT PSA_SALARIE,PSA_LIBELLE,PSA_PRENOM,PSA_LIBELLEEMPLOI,'+
               'PFI_SALARIE,PFI_LIBELLE,PFI_MOTIFINSCFOR,PFI_NIVPRIORITE,PFI_CODESTAGE,PFI_AUTRECOUT,PFI_NATUREFORM,PFI_RESPONSFOR,PFI_NBINSC,PFI_DUREESTAGE,PFI_ETATINSCFOR,PFI_MOTIFETATINSC '+
               'FROM DEPORTSAL LEFT JOIN INSCFORMATION ON (PSE_SALARIE=PFI_SALARIE AND PFI_MILLESIME="'+Millesime+'" AND PFI_CODESTAGE<>"--CURSUS--")'+
                             ' LEFT JOIN SALARIES ON (PSE_SALARIE=PSA_SALARIE) '+
               'WHERE (PSA_DATESORTIE IS NULL OR PSA_DATESORTIE>="'+UsdateTime(DateSortieSal)+'" OR PSA_DATESORTIE<="'+UsDateTime(IDate1900)+'")'+Where+
               ' ORDER BY PSA_LIBELLE',True);
  //PT1 - Fin

  TFor := Tob.Create('Table',Nil,-1);
  TFor.LoadDetailDB('Table','','',Q,False);
  Ferme(Q);
  //PT3 - Début
  // On fusionne les deux listes
  If TForN <> Nil Then
  Begin
     For i:= TForN.Detail.Count-1 DownTo 0 Do
          TForN.Detail[i].ChangeParent(TFor,0,False);
     FreeAndNil(TForN);
  End;
  //PT3 - Fin

  // Si pas de données : Grille vide
  If TFor.Detail.Count = 0 then
     Grille.Enabled := False
  Else
     Grille.Enabled := True;

  TGrille := Tob.Create('Table',Nil,-1);
  TSomme  := Tob.Create('sommes',Nil,-1);

  For i := 0 To TFor.Detail.Count - 1 Do
  begin
     If TFor.Detail[i].GetValue('PSA_SALARIE') <> Salarie Then    //PT3
     Begin
       //Insertion du compte-rendu du dernier salarié dans la grille
       If Salarie <> '' Then //PT3
       Begin
          TS := Tob.create('FilleSomme',TSomme,-1);
          TS.AddChampSupValeur('SALARIE',Salarie);
          TS.AddChampSupValeur('HEURES',NbHeures);
          TS.AddChampSupValeur('COUT',Cout);
          TS.AddChampSupValeur('NBINSC',NbInsc);
       End;

       NbHeures := 0;
       NbInsc   := 0;
       Cout     := 0;

       TG := Tob.create('Fille',TGrille,-1);
       Salarie := TFor.Detail[i].GetValue('PSA_SALARIE');
       If (Salarie = 'NULL') Then
       Begin
            TG.AddChampSupValeur('COL1','Inscriptions non nominatives');
            TG.AddChampSupValeur('COL2','');
       End
       Else
       Begin
            TG.AddChampSupValeur('COL1',TFor.Detail[i].GetValue('PSA_LIBELLE') +' '+TFor.Detail[i].GetValue('PSA_PRENOM'));
            TG.AddChampSupValeur('COL2',RechDom('PGLIBEMPLOI',TFor.Detail[i].GetValue('PSA_LIBELLEEMPLOI'),False));
       End;
       TG.AddChampSupValeur('COL2BIS', ''); //PT5
       TG.AddChampSupValeur('COL3','');
       TG.AddChampSupValeur('COL4','');
       TG.AddChampSupValeur('COL5','');
       TG.AddChampSupValeur('COL6','');
       TG.AddChampSupValeur('COL7','');
       TG.AddChampSupValeur('COL8','S');
       TG.AddChampSupValeur('COL9',Salarie);
       TG.AddChampSupValeur('COL10','');
       TG.AddChampSupValeur('COL11','');
       TG.AddChampSupValeur('SALARIE',Salarie);
     End;

     // Si PFI_LIBELLE non vide, il y a une inscription
     If TFor.Detail[i].GetValue('PFI_LIBELLE') <> '' Then //PT3
     Begin
            TG := Tob.create('Fille',TGrille,-1);
            TG.AddChampSupValeur('COL1',RechDom('PGSTAGEFORM',TFor.Detail[i].GetValue('PFI_CODESTAGE'),False));
            TG.AddChampSupValeur('COL2',RechDom('PGNATUREFORM',TFor.Detail[i].GetValue('PFI_NATUREFORM'),False));
            TG.AddChampSupValeur('COL2BIS', RechDom('PGINTERIMAIRES', TFor.Detail[i].GetValue('PFI_RESPONSFOR'),False)); //PT5
            TG.AddChampSupValeur('COL3',TFor.Detail[i].GetValue('PFI_NBINSC'));
            TG.AddChampSupValeur('COL4',TFor.Detail[i].GetValue('PFI_DUREESTAGE'));
            If TFor.Detail[i].GetValue('PFI_NATUREFORM') = '002' then
            begin
                 TG.AddChampSupValeur('COL5',TFor.Detail[i].GetValue('PFI_AUTRECOUT'));
                 Cout := Cout + TFor.Detail[i].GetValue('PFI_AUTRECOUT');
            end
            else TG.AddChampSupValeur('COL5',0);
            TG.AddChampSupValeur('COL6',RechDom('PGMOTIFINSCFORMATION',TFor.Detail[i].GetValue('PFI_MOTIFINSCFOR'),False));
            TG.AddChampSupValeur('COL7',RechDom('PGNIVPRIORITEFORM',TFor.Detail[i].GetValue('PFI_NIVPRIORITE'),False));
            TG.AddChampSupValeur('COL8','F');
            TG.AddChampSupValeur('COL9','');
            TG.AddChampSupValeur('COL10',RechDom('PGETATVALIDATION',TFor.Detail[i].GetValue('PFI_ETATINSCFOR'),False));
            TG.AddChampSupValeur('COL11',RechDom('PGMOTIFETATINSC',TFor.Detail[i].GetValue('PFI_MOTIFETATINSC'),False));
            TG.AddChampSupValeur('SALARIE','');

            NbHeures  := NbHeures  + TFor.Detail[i].GetValue('PFI_DUREESTAGE');
            NbInsc    := NbInsc    + TFor.Detail[i].GetValue('PFI_NBINSC');
            //PT5 - Début
            //TotInsc   := TotInsc   + TFor.Detail[i].GetValue('PFI_NBINSC');
            //TotHeures := TotHeures + TFor.Detail[i].GetValue('PFI_DUREESTAGE');
            //TotHono   := TotHono   + Cout;
            //PT5 - Fin
       end;
  end;

  //Insertion du compte-rendu du dernier salarié dans la grille
  If (TFor.Detail.Count>0) Then //PT3
  Begin
     TS := Tob.create('FilleSomme',TSomme,-1);
     TS.AddChampSupValeur('SALARIE',Salarie);
     TS.AddChampSupValeur('HEURES', NbHeures);
     TS.AddChampSupValeur('COUT',   Cout);
     TS.AddChampSupValeur('NBINSC', NbInsc);
  End;

  TFor.Free;

  For i := 0 to TGrille.Detail.Count - 1 do
  begin
       if TGrille.Detail[i].GetValue('COL8') = 'S' then
       begin
            TS := TSomme.FindFirst(['SALARIE'],[TGrille.Detail[i].GetValue('SALARIE')],False);
            If TS <> Nil then
            begin
                 TotSal := TotSal + 1;
                 TGrille.Detail[i].PutValue('COL3',TS.GetValue('NBINSC'));
                 TGrille.Detail[i].PutValue('COL4',TS.GetValue('HEURES'));
                 TGrille.Detail[i].PutValue('COL5',TS.GetValue('COUT'));
                 
	            //PT5 - Début
	            TotInsc   := TotInsc   + TS.GetValue('NBINSC');
	            TotHeures := TotHeures + TS.GetValue('HEURES');
	            TotHono   := TotHono   + TS.GetValue('COUT');
	            //PT5 - Fin
            end;
       end;
  end;
  TSomme.Free;

  If TGrille.Detail.Count > 0 Then   //PT3
  	If (GetCheckBoxState('MULTINIVEAU')=CbChecked) Then  //PT5
     TGrille.PutGridDetail(THGrid(GetControl('GRILLE')),False,False,'COL1;COL2;COL2BIS;COL3;COL4;COL5;COL6;COL7;COL8;COL9;COL10;COL11',False) 
    else
     TGrille.PutGridDetail(THGrid(GetControl('GRILLE')),False,False,'COL1;COL2;COL3;COL4;COL5;COL6;COL7;COL8;COL9;COL10;COL11',False);
  TGrille.Free;
  
  Grille.CellValues[0,1] := 'Formation';
  Grille.CellValues[1,1] := 'Nature';
  
  //PT5 - Début
  If (GetCheckBoxState('MULTINIVEAU')=CbChecked) Then 
  	Col := 8
  else
  	Col := 7;
  //PT5 -Fin
  For i := 2 to Grille.RowCount - 1 do
  begin
       If Grille.CellValues[Col,i] = 'S' then Grille.RowHeights[i] := 24 //PT5
       else Grille.RowHeights[i] := 15;
  end;
  Grille.DefaultColWidth := 100;

  //PT3 - Début
  Grille.Cells[0,0] := 'Salarié';
  Grille.Cells[1,0] := 'Libellé emploi';
  Col := 2;
  If NbCol = 12 Then Begin Grille.Cells[Col,0] := 'Responsable'; Col := Col + 1; End; //PT5
  Grille.Cells[Col,0] := 'Inscrits';			Col := Col + 1;
  Grille.Cells[Col,0] := 'Heures';              Col := Col + 1;
  Grille.Cells[Col,0] := 'Honoraires';          Col := Col + 1;
  Grille.Cells[Col,0] := 'Priorité';            Col := Col + 1;
  Grille.Cells[Col,0] := 'Prévision';           Col := Col + 1;
  Grille.Cells[Col,0] := '';                    Col := Col + 1;
  Grille.Cells[Col,0] := '';                    Col := Col + 1;
  Grille.Cells[Col,0] := 'Etat';                Col := Col + 1;
  Grille.Cells[Col,0]:= 'Motif';

  Grille.ColWidths[0] := 200;
  Grille.ColWidths[1] := 150;
  Col := 2;
  If NbCol = 12 Then Begin Grille.ColWidths[2] := 100; Col := Col + 1; End; //PT5
  Grille.ColWidths[Col] := 50;					Col := Col + 1;
  Grille.ColWidths[Col] := 50;					Col := Col + 1;
  Grille.ColWidths[Col] := 50;                  Col := Col + 1;
  Grille.ColWidths[Col] := 50;                  Col := Col + 1;
  Grille.ColWidths[Col] := 50;                  Col := Col + 1;
  Grille.ColWidths[Col] := -1;                  Col := Col + 1;
  Grille.ColWidths[Col] := -1;                  Col := Col + 1;
  Grille.ColWidths[Col] := 50;                  Col := Col + 1;
  Grille.ColWidths[Col] := 70;

  Col := 2;
  If NbCol = 12 Then Col := Col + 1; //PT5
  Grille.ColFormats[Col] := '# ##0';			Col := Col + 1;
  Grille.ColFormats[Col] := '# ##0.00';         Col := Col + 1;
  Grille.ColFormats[Col] := '# ##0';            Col := Col + 1;
  Grille.ColFormats[Col] := '';                 Col := Col + 1;
  Grille.ColFormats[Col] := '';

  Col := 2;
  If NbCol = 12 Then Begin Grille.ColAligns[Col] := taLeftJustify; Col := Col + 1; End; //PT5
  Grille.ColAligns[Col] := taRightJustify;     	Col := Col + 1;
  Grille.ColAligns[Col] := taRightJustify;      Col := Col + 1;
  Grille.ColAligns[Col] := taRightJustify;      Col := Col + 1;
  Grille.ColAligns[Col] := taLeftJustify;       Col := Col + 1;
  Grille.ColAligns[Col] := taLeftJustify;
  //PT3 - Fin
  
  TFVierge(Ecran).HMTrad.ResizeGridColumns(Grille) ;
  SetControlText('TOTINSC',floatToStr(TotInsc));
  SetControlText('TOTHEURES',floatToStr(TotHeures));
  SetControlText('TOTHONO',floatToStr(Tothono));
end;



Procedure TOF_PGTESTFORMATION.AffichageParFormation;
var Q : TQuery;
    TG,TGrille,TFor : Tob;
    i : Integer;
    Formation,Nature,Where,SQL : String; //PT1
    NbHeures,Cout : Double;
    NbInsc : Integer;
    TSomme,TS : Tob;
    TotHeures,TotInsc,Tothono : Double;
    NbCol, Col : Integer; //PT5
begin
  TotInsc   := 0;
  TotHeures := 0;
  TotHono   := 0;
  NbInsc    := 0;
  Cout      := 0;
  NbHeures  := 0;

  Grille.Visible := True;
  For i := 0 to Grille.RowCount - 1 do Grille.Rows[i].Clear;
  Grille.RowCount  := 3;    //PT3
  Grille.FixedRows := 2;

  //PT5 - Début
  If (GetCheckBoxState('MULTINIVEAU')=CbChecked) Then
    NbCol := 12
  Else
    NbCol := 11; //PT3
  Grille.ColCount := NbCol;
  //PT5 - Fin

  //PT1 - Début
  // Critères de sélection
  If GetControlText('FAVORIS') = 'X' Then
    Where := ' AND PST_CODESTAGE IN (SELECT PCC_CODESTAGE FROM CURSUSSTAGE WHERE PCC_CURSUS="---" AND PCC_LIBELLE="'+V_PGI.UserSalarie+'") '
  Else
    Where := ' AND PST_CODESTAGE NOT IN (SELECT PCC_CODESTAGE FROM CURSUSSTAGE WHERE PCC_CURSUS="---" AND PCC_LIBELLE="'+V_PGI.UserSalarie+'") ';
  //PT1 - Fin
  //PT3 - Début
  SQL := '';
  If GetControlText('PSA_SALARIE') <> '' Then SQL := SQL + ' AND PFI_SALARIE="'+GetControlText('PSA_SALARIE')+'"';
  If GetControlText('PSA_LIBELLEEMPLOI') <> '' Then SQL := SQL + ' AND PFI_LIBEMPLOIFOR="'+GetControlText('PSA_LIBELLEEMPLOI')+'"';
  //PT1 - Début
  SQL := SQL + ' AND ' + AdaptByRespEmanager (TypeUtilisat,'PFI',Utilisateur,(GetCheckBoxState('MULTINIVEAU')=CbChecked)); //PT5

  Q := OpenSQL('SELECT PST_CODESTAGE,PST_MILLESIME,PST_NATUREFORM,PST_LIBELLE,PST_LIBELLE1,'+
               'PFI_SALARIE,PFI_MOTIFINSCFOR,PFI_NIVPRIORITE,PFI_CODESTAGE,PFI_AUTRECOUT,PFI_NATUREFORM,PFI_RESPONSFOR,'+ //PT5
               'PFI_NBINSC,PFI_DUREESTAGE,PFI_LIBELLE,PFI_LIBEMPLOIFOR,PFI_ETATINSCFOR,PFI_MOTIFETATINSC '+
               'FROM STAGE LEFT JOIN INSCFORMATION ON PST_CODESTAGE=PFI_CODESTAGE '+
               ' AND PFI_MILLESIME="'+Millesime+'" '+'AND '+
               SQL+
               ' WHERE '+
               'PST_ACTIF="X" AND PST_MILLESIME="0000" AND PST_FORMATION3="'+CYCLE_EN_COURS_EM+'"'+Where+
               ' ORDER BY PST_CODESTAGE,PFI_LIBELLE',True);
  //PT3- Fin
  //PT1 - Fin

  TFor := Tob.Create('Table',Nil,-1);
  TFor.LoadDetailDB('Table','','',Q,False);
  Ferme(Q);

  If TFor.Detail.Count = 0 then //PT3
     Grille.Enabled := False
  Else
     Grille.Enabled := True;

  TGrille := Tob.Create('Table',Nil,-1);
  TSomme := Tob.Create('sommes',Nil,-1);

  For i := 0 to TFor.Detail.Count - 1 Do
  Begin
     If TFor.Detail[i].GetValue('PST_CODESTAGE') <> Formation Then    //PT3
     Begin

       //Insertion du compte-rendu du dernier stage dans la grille
       If Formation <> '' Then //PT3
       Begin
            TS := Tob.create('FilleSomme',TSomme,-1);
            TS.AddChampSupValeur('STAGE',Formation);
            TS.AddChampSupValeur('HEURES',NbHeures);
            TS.AddChampSupValeur('COUT',Cout);
            TS.AddChampSupValeur('NBINSC',NbInsc);
       End;

       Formation := TFor.Detail[i].GetValue('PST_CODESTAGE');
       Nature    := TFor.Detail[i].GetValue('PST_NATUREFORM');
       NbHeures  := 0;
       NbInsc    := 0;
       Cout      := 0;
       
       TG := Tob.create('Fille',TGrille,-1);
       TG.AddChampSupValeur('COL1',TFor.Detail[i].GetValue('PST_LIBELLE') + TFor.Detail[i].GetValue('PST_LIBELLE1'));
       TG.AddChampSupValeur('COL2',RechDom('PGNATUREFORM',Nature,False));
       TG.AddChampSupValeur('COL2BIS', ''); //PT5
       TG.AddChampSupValeur('COL3','');
       TG.AddChampSupValeur('COL4','');
       TG.AddChampSupValeur('COL5','');
       TG.AddChampSupValeur('COL6','');
       TG.AddChampSupValeur('COL7','');
       TG.AddChampSupValeur('COL8','F');
       TG.AddChampSupValeur('COL9',Formation);
       TG.AddChampSupValeur('COL10','');
       TG.AddChampSupValeur('COL11','');
       TG.AddChampSupValeur('STAGE',Formation);
     End;

     // Si PFI_LIBELLE non vide, il y a une inscription
     If TFor.Detail[i].GetValue('PFI_LIBELLE') <> '' Then //PT3
     Begin
            TG := Tob.create('Fille',TGrille,-1);
            TG.AddChampSupValeur('COL1',TFor.Detail[i].GetValue('PFI_LIBELLE'));
            TG.AddChampSupValeur('COL2',RechDom('PGLIBEMPLOI',TFor.Detail[i].GetValue('PFI_LIBEMPLOIFOR'),False));
            TG.AddChampSupValeur('COL2BIS', RechDom('PGINTERIMAIRES', TFor.Detail[i].GetValue('PFI_RESPONSFOR'),False)); //PT5
            TG.AddChampSupValeur('COL3',TFor.Detail[i].GetValue('PFI_NBINSC'));
            TG.AddChampSupValeur('COL4',TFor.Detail[i].GetValue('PFI_DUREESTAGE'));
            If TFor.Detail[i].GetValue('PFI_NATUREFORM') = '002' then
            begin
                 TG.AddChampSupValeur('COL5',TFor.Detail[i].GetValue('PFI_AUTRECOUT'));
                 Cout := Cout + TFor.Detail[i].GetValue('PFI_AUTRECOUT');
            end
            else TG.AddChampSupValeur('COL5',0);
            TG.AddChampSupValeur('COL6',RechDom('PGMOTIFINSCFORMATION',TFor.Detail[i].GetValue('PFI_MOTIFINSCFOR'),False));
            TG.AddChampSupValeur('COL7',RechDom('PGNIVPRIORITEFORM',TFor.Detail[i].GetValue('PFI_NIVPRIORITE'),False));
            TG.AddChampSupValeur('COL8','S');
            TG.AddChampSupValeur('COL9','');
            TG.AddChampSupValeur('COL10',RechDom('PGETATVALIDATION',TFor.Detail[i].GetValue('PFI_ETATINSCFOR'),False));
            TG.AddChampSupValeur('COL11',RechDom('PGMOTIFETATINSC',TFor.Detail[i].GetValue('PFI_MOTIFETATINSC'),False));
            TG.AddChampSupValeur('STAGE','');
            
            NbHeures  := NbHeures  + TFor.Detail[i].GetValue('PFI_DUREESTAGE');
            NbInsc    := NbInsc    + TFor.Detail[i].GetValue('PFI_NBINSC');
            //PT5 - Début
            //TotInsc   := TotInsc   + TFor.Detail[i].GetValue('PFI_NBINSC');
            //TotHeures := TotHeures + TFor.Detail[i].GetValue('PFI_DUREESTAGE');
            //TotHono   := TotHono   + Cout;
            //PT5 - Fin
       end;
  end;
  //PT3 - Début
  //Insertion du compte-rendu du dernier stage dans la grille
  If (TFor.Detail.Count>0) {And (NbInsc > 0) }Then
  Begin
     TS := Tob.create('FilleSomme',TSomme,-1);
     TS.AddChampSupValeur('STAGE',Formation);
     TS.AddChampSupValeur('HEURES',NbHeures);
     TS.AddChampSupValeur('COUT',Cout);
     TS.AddChampSupValeur('NBINSC',NbInsc);
  End;
  //PT3 - Fin

  TFor.Free;
  For i := 0 to TGrille.Detail.Count - 1 do
  begin
       TS := TSomme.FindFirst(['STAGE'],[TGrille.Detail[i].GetValue('STAGE')],False);
       If TS <> Nil then
       begin
            TGrille.Detail[i].PutValue('COL3',TS.GetValue('NBINSC'));
            TGrille.Detail[i].PutValue('COL4',TS.GetValue('HEURES'));
            TGrille.Detail[i].PutValue('COL5',TS.GetValue('COUT'));
            
            //PT5 - Début
            TotInsc   := TotInsc   + TS.GetValue('NBINSC');
            TotHeures := TotHeures + TS.GetValue('HEURES');
            TotHono   := TotHono   + TS.GetValue('COUT');
            //PT5 - Fin
       end;
  end;
  TSomme.Free;

  //PT3 - Début
  // Mise en forme de la grille
  Grille.Cells[0,0] := 'Formation';
  Grille.Cells[1,0] := 'Nature';
  Col := 2;
  If NbCol = 12 Then Begin Grille.Cells[Col,0] := 'Responsable'; Col := Col + 1; End; //PT5
  Grille.Cells[Col,0] := 'Inscrits';			Col := Col + 1;
  Grille.Cells[Col,0] := 'Heures';              Col := Col + 1;
  Grille.Cells[Col,0] := 'Honoraires';          Col := Col + 1;
  Grille.Cells[Col,0] := 'Priorité';            Col := Col + 1;
  Grille.Cells[Col,0] := 'Prévision';           Col := Col + 1;
  Grille.Cells[Col,0] := '';                    Col := Col + 1;
  Grille.Cells[Col,0] := '';                    Col := Col + 1;
  Grille.Cells[Col,0]:= 'Etat';                 Col := Col + 1;
  Grille.Cells[Col,0]:= 'Motif';

  Grille.DefaultColWidth := 100;

  Grille.ColWidths[0] := 200;
  Grille.ColWidths[1] := 150;
  Col := 2;
  If NbCol = 12 Then Begin Grille.ColWidths[2] := 100; Col := Col + 1; End; //PT5
  Grille.ColWidths[Col] := 50;					Col := Col + 1;
  Grille.ColWidths[Col] := 50;                  Col := Col + 1;
  Grille.ColWidths[Col] := 50;                  Col := Col + 1;
  Grille.ColWidths[Col] := 50;                  Col := Col + 1;
  Grille.ColWidths[Col] := 50;                  Col := Col + 1;
  Grille.ColWidths[Col] := -1;                  Col := Col + 1;
  Grille.ColWidths[Col] := -1;                  Col := Col + 1;
  Grille.ColWidths[Col]:= 50;                   Col := Col + 1;
  Grille.ColWidths[Col]:= 80;

  Col := 2;
  If NbCol = 12 Then Col := Col + 1; //PT5
  Grille.ColFormats[Col] := '# ##0';			Col := Col + 1;
  Grille.ColFormats[Col] := '# ##0.00';         Col := Col + 1;
  Grille.ColFormats[Col] := '# ##0';            Col := Col + 1;
  Grille.ColFormats[Col] := '';                 Col := Col + 1;
  Grille.ColFormats[Col] := '';
  
  Col := 2;
  If NbCol = 12 Then Begin Grille.ColAligns[Col] := taLeftJustify; Col := Col + 1; End; //PT5
  Grille.ColAligns[Col] := taRightJustify;		Col := Col + 1;
  Grille.ColAligns[Col] := taRightJustify;      Col := Col + 1;
  Grille.ColAligns[Col] := taRightJustify;      Col := Col + 1;
  Grille.ColAligns[Col] := taLeftJustify;       Col := Col + 1;
  Grille.ColAligns[Col] := taLeftJustify;
  //PT3 - Fin

  If TGrille.Detail.Count > 0 Then  //PT3
  	If (GetCheckBoxState('MULTINIVEAU')=CbChecked) Then  //PT5
     TGrille.PutGridDetail(THGrid(GetControl('GRILLE')),False,False,'COL1;COL2;COL2BIS;COL3;COL4;COL5;COL6;COL7;COL8;COL9;COL10;COL11',False) 
    else
     TGrille.PutGridDetail(THGrid(GetControl('GRILLE')),False,False,'COL1;COL2;COL3;COL4;COL5;COL6;COL7;COL8;COL9;COL10;COL11',False);

  TGrille.Free;
  Grille.CellValues[0,1] := 'Salarié';
  Grille.CellValues[1,1] := 'Libellé emploi';

  //PT5 - Début
  If (GetCheckBoxState('MULTINIVEAU')=CbChecked) Then 
  	Col := 8
  else
  	Col := 7;
  //PT5 -Fin
  For i := 2 to Grille.RowCount - 1 do
  begin
   		If Grille.CellValues[Col,i] = 'F' then Grille.RowHeights[i] := 24 //PT5
   		else Grille.RowHeights[i] := 15;
  end;

  TFVierge(Ecran).HMTrad.ResizeGridColumns(Grille) ;
  SetControlText('TOTINSC',floatToStr(TotInsc));
  SetControlText('TOTHEURES',floatToStr(TotHeures));
  SetControlText('TOTHONO',floatToStr(Tothono));
end;

//PT1 - Début
{********************************************************
                  AFFICHAGE PAR CURSUS
********************************************************}
procedure TOF_PGTESTFORMATION.AffichageParCursus;
var Q : TQuery;
    TG,TGrille,TCur : Tob;
    i : Integer;
    Cursus,SQL,Where : String;
    NbHeures,Cout : Double;
    NbInsc : Integer;
    TSomme,TS : Tob;
    TotHeures,TotInsc,Tothono : Double;
    NbCol, Col : Integer; //PT5
begin
  TotInsc   := 0;
  TotHeures := 0;
  TotHono   := 0;
  NbInsc    := 0;
  Cout      := 0;
  NbHeures  := 0;

  Grille.Visible := True;
  For i := 0 to Grille.RowCount - 1 do Grille.Rows[i].Clear;
  Grille.RowCount  := 3;
  Grille.FixedRows := 2;
  
  //PT5 - Début
  If (GetCheckBoxState('MULTINIVEAU')=CbChecked) Then
    NbCol := 12
  Else
    NbCol := 11; //PT3
  Grille.ColCount := NbCol;
  //PT5 - Fin

  // Critères de sélection
  //PT4 - Début
  If GetControlText('FAVORIS') = 'X' Then
    Where := ' AND PCU_CURSUS IN (SELECT PCC_CODESTAGE FROM CURSUSSTAGE WHERE PCC_CURSUS="-C-" AND PCC_LIBELLE="'+V_PGI.UserSalarie+'") '
  Else
    Where := ' AND PCU_CURSUS NOT IN (SELECT PCC_CODESTAGE FROM CURSUSSTAGE WHERE PCC_CURSUS="-C-" AND PCC_LIBELLE="'+V_PGI.UserSalarie+'") ';
  //PT4 - Fin
  //PT3 - Début
  SQL := '';
  If GetControlText('PSA_SALARIE') <> '' Then SQL := SQL + ' AND PFI_SALARIE="'+GetControlText('PSA_SALARIE')+'"';
  If GetControlText('PSA_LIBELLEEMPLOI') <> '' Then SQL := SQL + ' AND PFI_LIBEMPLOIFOR="'+GetControlText('PSA_LIBELLEEMPLOI')+'"';
  //PT1 - Début
  SQL := SQL + ' AND ' + AdaptByRespEmanager (TypeUtilisat,'PFI',Utilisateur,(GetCheckBoxState('MULTINIVEAU')=CbChecked)); //PT5

  // Liste des cursus
  Q := OpenSQL('SELECT PCU_CURSUS, PCU_LIBELLE, PCU_RANGCURSUS,'+
               'PFI_MOTIFINSCFOR,PFI_NIVPRIORITE,PFI_SALARIE,PFI_CODESTAGE,PFI_AUTRECOUT,PFI_NATUREFORM,PFI_RESPONSFOR,PFI_NBINSC,'+
               'PFI_DUREESTAGE,PFI_LIBELLE,PFI_LIBEMPLOIFOR,PFI_ETATINSCFOR,PFI_MOTIFETATINSC '+
               'FROM CURSUS LEFT JOIN INSCFORMATION ON PCU_CURSUS=PFI_CURSUS AND PFI_MILLESIME="'+Millesime+'" AND PFI_CODESTAGE="--CURSUS--"'+
               SQL+
               ' WHERE PCU_RANGCURSUS=0 AND PCU_INCLUSCAT="X" '+Where+
               'ORDER BY PCU_CURSUS,PFI_LIBELLE', True);
  //PT3 - Fin
  Tcur := Tob.Create('Table',Nil,-1);
  Tcur.LoadDetailDB('Table','','',Q,False);
  Ferme(Q);

  // Si pas de données : Grille vide
  If Tcur.Detail.Count = 0 then
     Grille.Enabled := False
  Else
     Grille.Enabled := True;

  // Si données présentes : Liste des inscrits par cursus
  TGrille := Tob.Create('Table',Nil,-1);
  TSomme  := Tob.Create('sommes',Nil,-1);

  For i := 0 to Tcur.Detail.Count - 1 do
  begin
     If TCur.Detail[i].GetValue('PCU_CURSUS') <> Cursus Then    //PT3
     Begin
       //Insertion du compte-rendu du dernier cursus dans la grille
       If Cursus <> '' Then //PT3
       Begin
          TS := Tob.create('FilleSomme',TSomme,-1);
          TS.AddChampSupValeur('CURSUS',Cursus);
          TS.AddChampSupValeur('HEURES',NbHeures);
          TS.AddChampSupValeur('COUT',Cout);
          TS.AddChampSupValeur('NBINSC',NbInsc);
       End;

       Cursus   := TCur.Detail[i].GetValue('PCU_CURSUS');
       NbHeures := 0;
       NbInsc   := 0;
       Cout     := 0;

       TG := Tob.create('Fille',TGrille,-1);
       TG.AddChampSupValeur('COL1',TCur.Detail[i].GetValue('PCU_LIBELLE'));
       TG.AddChampSupValeur('COL2','');
       TG.AddChampSupValeur('COL2BIS', ''); //PT5
       TG.AddChampSupValeur('COL3','');
       TG.AddChampSupValeur('COL4','');
       TG.AddChampSupValeur('COL5','');
       TG.AddChampSupValeur('COL6','');
       TG.AddChampSupValeur('COL7','');
       TG.AddChampSupValeur('COL8','C');
       TG.AddChampSupValeur('COL9',Cursus);
       TG.AddChampSupValeur('COL10','');
       TG.AddChampSupValeur('COL11','');
       TG.AddChampSupValeur('CURSUS',Cursus);
     End;

     // Si PFI_LIBELLE non vide, il y a une inscription
     If TCur.Detail[i].GetValue('PFI_LIBELLE') <> '' Then //PT3
     Begin
            TG := Tob.create('Fille',TGrille,-1);
            TG.AddChampSupValeur('COL1',TCur.Detail[i].GetValue('PFI_LIBELLE'));
            TG.AddChampSupValeur('COL2',RechDom('PGLIBEMPLOI',TCur.Detail[i].GetValue('PFI_LIBEMPLOIFOR'),False));
            TG.AddChampSupValeur('COL2BIS', RechDom('PGINTERIMAIRES', TCur.Detail[i].GetValue('PFI_RESPONSFOR'),False)); //PT5
            TG.AddChampSupValeur('COL3',TCur.Detail[i].GetValue('PFI_NBINSC'));
            TG.AddChampSupValeur('COL4',TCur.Detail[i].GetValue('PFI_DUREESTAGE'));
            TG.AddChampSupValeur('COL5',0);
            TG.AddChampSupValeur('COL6',RechDom('PGMOTIFINSCFORMATION',TCur.Detail[i].GetValue('PFI_MOTIFINSCFOR'),False));
            TG.AddChampSupValeur('COL7',RechDom('PGNIVPRIORITEFORM',TCur.Detail[i].GetValue('PFI_NIVPRIORITE'),False));
            TG.AddChampSupValeur('COL8','S');
            TG.AddChampSupValeur('COL9','');
            TG.AddChampSupValeur('COL10',RechDom('PGETATVALIDATION',TCur.Detail[i].GetValue('PFI_ETATINSCFOR'),False));
            TG.AddChampSupValeur('COL11',RechDom('PGMOTIFETATINSC',TCur.Detail[i].GetValue('PFI_MOTIFETATINSC'),False));
            TG.AddChampSupValeur('CURSUS','');

            NbHeures  := NbHeures  + TCur.Detail[i].GetValue('PFI_DUREESTAGE');
            NbInsc    := NbInsc    + TCur.Detail[i].GetValue('PFI_NBINSC');
            //PT5 - Début
            //TotInsc   := TotInsc   + TCur.Detail[i].GetValue('PFI_NBINSC');
            //TotHeures := TotHeures + TCur.Detail[i].GetValue('PFI_DUREESTAGE');
            //TotHono   := TotHono   + Cout;
            //PT5 - Fin
       end;
  end;

  //PT3 - Début
  //Insertion du compte-rendu du dernier stage dans la grille
  If (TCur.Detail.Count>0) {And (NbInsc > 0) }Then
  Begin
       TS := Tob.create('FilleSomme',TSomme,-1);
       TS.AddChampSupValeur('CURSUS',Cursus);
       TS.AddChampSupValeur('HEURES',NbHeures);
       TS.AddChampSupValeur('COUT',Cout);
       TS.AddChampSupValeur('NBINSC',NbInsc);
  End;
  Tcur.Free;

  // Totaux
  For i := 0 to TGrille.Detail.Count - 1 do
  begin
       TS := TSomme.FindFirst(['CURSUS'],[TGrille.Detail[i].GetValue('CURSUS')],False);
       If TS <> Nil then
       begin
            TGrille.Detail[i].PutValue('COL3',TS.GetValue('NBINSC'));
            TGrille.Detail[i].PutValue('COL4',TS.GetValue('HEURES'));
            TGrille.Detail[i].PutValue('COL5',TS.GetValue('COUT'));
            
            //PT5 - Début
            TotInsc   := TotInsc   + TS.GetValue('NBINSC');
            TotHeures := TotHeures + TS.GetValue('HEURES');
            TotHono   := TotHono   + TS.GetValue('COUT');
            //PT5 - Fin
       end;
  end;
  TSomme.Free;

  // Insertion des infos de la TOB en grille
  If TGrille.Detail.Count > 0 Then  //PT3
  	If (GetCheckBoxState('MULTINIVEAU')=CbChecked) Then  //PT5
     TGrille.PutGridDetail(THGrid(GetControl('GRILLE')),False,False,'COL1;COL2;COL2BIS;COL3;COL4;COL5;COL6;COL7;COL8;COL9;COL10;COL11',False) 
    else
     TGrille.PutGridDetail(THGrid(GetControl('GRILLE')),False,False,'COL1;COL2;COL3;COL4;COL5;COL6;COL7;COL8;COL9;COL10;COL11',False);
  TGrille.Free;

  // Format de la grille
  Grille.CellValues[0,1] := 'Salarié';
  Grille.CellValues[1,1] := 'Libellé emploi';
  
  //PT5 - Début
  If (GetCheckBoxState('MULTINIVEAU')=CbChecked) Then 
  	Col := 8
  else
  	Col := 7;
  //PT5 -Fin
  
  For i := 2 to Grille.RowCount - 1 do
  begin
       If Grille.CellValues[Col,i] = 'C' then Grille.RowHeights[i] := 24
       else Grille.RowHeights[i] := 15;
  end;
  Grille.DefaultColWidth := 100;

  //PT3 - Début
  // Mise en forme de la grille
  Grille.Cells[0,0] := 'Cursus';
  Grille.Cells[1,0] := '';
  Col := 2;
  If NbCol = 12 Then Begin Grille.Cells[Col,0] := 'Responsable'; Col := Col + 1; End; //PT5
  Grille.Cells[Col,0] := 'Inscrits';			Col := Col + 1;
  Grille.Cells[Col,0] := 'Heures';              Col := Col + 1;
  Grille.Cells[Col,0] := 'Honoraires';          Col := Col + 1;
  Grille.Cells[Col,0] := 'Priorité';            Col := Col + 1;
  Grille.Cells[Col,0] := 'Prévision';           Col := Col + 1;
  Grille.Cells[Col,0] := '';                    Col := Col + 1;
  Grille.Cells[Col,0] := '';                    Col := Col + 1;
  Grille.Cells[Col,0] := 'Etat';                Col := Col + 1;
  Grille.Cells[Col,0]:= 'Motif';

  Grille.ColWidths[0] := 200;
  Grille.ColWidths[1] := 150;
  Col := 2;
  If NbCol = 12 Then Begin Grille.ColWidths[2] := 100; Col := Col + 1; End; //PT5
  Grille.ColWidths[Col] := 50;					Col := Col + 1;
  Grille.ColWidths[Col] := 50;                  Col := Col + 1;
  Grille.ColWidths[Col] := 50;                  Col := Col + 1;
  Grille.ColWidths[Col] := 50;                  Col := Col + 1;
  Grille.ColWidths[Col] := 50;                  Col := Col + 1;
  Grille.ColWidths[Col] := -1;                  Col := Col + 1;
  Grille.ColWidths[Col] := -1;                  Col := Col + 1;
  Grille.ColWidths[Col] := 50;                  Col := Col + 1;
  Grille.ColWidths[Col] := 70;

  Col := 2;
  If NbCol = 12 Then Col := Col + 1; //PT5
  Grille.ColFormats[Col] := '# ##0';			Col := Col + 1;
  Grille.ColFormats[Col] := '# ##0.00';         Col := Col + 1;
  Grille.ColFormats[Col] := '# ##0';            Col := Col + 1;
  Grille.ColFormats[Col] := '';                 Col := Col + 1;
  Grille.ColFormats[Col] := '';
  
  Col := 2;
  If NbCol = 12 Then Begin Grille.ColAligns[Col] := taLeftJustify; Col := Col + 1; End; //PT5
  Grille.ColAligns[Col] := taRightJustify;		Col := Col + 1;
  Grille.ColAligns[Col] := taRightJustify;      Col := Col + 1;
  Grille.ColAligns[Col] := taRightJustify;      Col := Col + 1;
  Grille.ColAligns[Col] := taLeftJustify;       Col := Col + 1;
  Grille.ColAligns[Col] := taLeftJustify;
  //PT3 - Fin
  
  TFVierge(Ecran).HMTrad.ResizeGridColumns(Grille) ;
  SetControlText('TOTINSC',floatToStr(TotInsc));
  SetControlText('TOTHEURES',floatToStr(TotHeures));
  SetControlText('TOTHONO',floatToStr(Tothono));
end;
//PT1 - Fin

(*procedure TOF_PGTESTFORMATION.AffichageParRespFormation(Sender : Tobject);    //PT5
var Q : TQuery;
    TG,TGrille,TFor,TForN : Tob;
    i : Integer;
    rESP : String;
    NbHeures,Cout : Double;
    NbInsc : Integer;
    TSomme,TS : Tob;
    Where : String;
    TotHeures,TotInsc,Tothono : Double;
begin
  TotInsc   := 0;
  TotHeures := 0;
  TotHono   := 0;
  NbInsc    := 0;
  Cout      := 0;
  NbHeures  := 0;

  //PT3 - Début
  SetControlProperty('PSA_LIBELLEEMPLOI','Plus', ' AND CC_CODE IN (SELECT PSA_LIBELLEEMPLOI FROM SALARIES LEFT JOIN DEPORTSAL ON PSA_SALARIE=PSE_SALARIE '+
  'WHERE (PSA_DATESORTIE IS NULL OR PSA_DATESORTIE>="'+UsdateTime(V_PGI.DateEntree)+'" OR PSA_DATESORTIE<="'+UsDateTime(IDate1900)+'") AND '+
  AdaptByRespEmanager (TypeUtilisat,'PSE',Utilisateur,(GetControlText('MULTINIVEAU')='X'))+')');
  //PT3 - Fin

  If GetCheckBoxState('MULTINIVEAU') <> CbChecked then
  begin
       SetControlChecked('RESPONSFOR',False);
       ChangeRegroupement(Nil);
       Exit;
  end;
  SetControlChecked('RESPONSFOR',true);;

  For i := 0 to Grille.RowCount - 1 Do Grille.Rows[i].Clear;
  Grille.RowCount := 3;
  Grille.ColCount := 13; //PT3

  //PT3 - Début
  // Cas particulier des inscriptions non nominatives
  // ------------------------------------------------
  If GetControlText('PSA_SALARIE') = '' Then
  Begin
       Where := '';
       If GetControlText('PSA_LIBELLEEMPLOI') <> '' Then Where := Where + ' AND PFI_LIBEMPLOIFOR="'+GetControlText('PSA_LIBELLEEMPLOI')+'"';
       Where := Where + ' AND ' + AdaptByRespEmanager (TypeUtilisat,'PFI',Utilisateur,True);

       Q := OpenSQL('SELECT PFI_RESPONSFOR AS PSI_INTERIMAIRE,PSA_LIBELLE AS PSI_LIBELLE,PSA_PRENOM AS PSI_PRENOM, "" AS PFI_SALARIE,'+
                    'PFI_LIBELLE,PFI_LIBEMPLOIFOR,PFI_MOTIFINSCFOR,PFI_NIVPRIORITE,PFI_CODESTAGE,PFI_AUTRECOUT,PFI_NATUREFORM,PFI_NBINSC,PFI_DUREESTAGE,PFI_ETATINSCFOR,PFI_MOTIFETATINSC '+
                    'FROM INSCFORMATION LEFT JOIN SALARIES ON (PFI_RESPONSFOR=PSA_SALARIE)'+
                    'WHERE PFI_MILLESIME="'+Millesime+'" AND PFI_CODESTAGE<>"--CURSUS--" AND PFI_SALARIE=""'+Where+
                    ' ORDER BY PSI_INTERIMAIRE',True);

       TForN := Tob.Create('Table',Nil,-1);
       TForN.LoadDetailDB('Table','','',Q,False);
       Ferme(Q);
  end;
  //PT3 - Fin

  // Autres inscriptions
  // -------------------
  // Critères de sélection
  Where := '';
  If GetControlText('PSA_SALARIE') <> '' then Where := Where + ' AND PFI_SALARIE="'+GetControlText('PSA_SALARIE')+'"';
  If GetControlText('PSA_LIBELLEEMPLOI') <> '' then Where := Where + ' AND PFI_LIBEMPLOIFOR="'+GetControlText('PSA_LIBELLEEMPLOI')+'"';
  //If GetControlText('PSA_ETABLISSEMENT') <> '' then Where := Where + ' AND PFI_ETABLISSEMENT="'+GetControlText('PSA_ETABLISSEMENT')+'"'; //PT1
  //PT1 - Début
  Where := Where + ' AND ' + AdaptByRespEmanager (TypeUtilisat,'PFI',Utilisateur,True);
  //PT1 - Fin

  //PT3 - Début
  Q := OpenSQL('SELECT PSI_INTERIMAIRE,PSI_LIBELLE,PSI_PRENOM,'+
//               'PFI_MOTIFINSCFOR,PFI_NIVPRIORITE,PFI_SALARIE,PFI_LIBEMPLOIFOR,PFI_CODESTAGE,PFI_NATUREFORM,SUM (PFI_NBINSC) NBINSC,SUM (PFI_DUREESTAGE) DUREESTAGE,PFI_AUTRECOUT,PFI_ETATINSCFOR,PFI_MOTIFETATINSC '+
               'PFI_SALARIE,PFI_LIBELLE,PFI_LIBEMPLOIFOR,PFI_MOTIFINSCFOR,PFI_NIVPRIORITE,PFI_CODESTAGE,PFI_AUTRECOUT,PFI_NATUREFORM,PFI_NBINSC,PFI_DUREESTAGE,PFI_ETATINSCFOR,PFI_MOTIFETATINSC '+
               'FROM INTERIMAIRES '+
               'LEFT JOIN INSCFORMATION ON (PFI_RESPONSFOR=PSI_INTERIMAIRE AND PFI_MILLESIME="'+Millesime+'" AND PFI_CODESTAGE<>"--CURSUS--") '+ //PT1
               'WHERE '+Where+
//               ' GROUP BY PSI_INTERIMAIRE,PSI_LIBELLE,PSI_PRENOM,PFI_MOTIFINSCFOR,PFI_NIVPRIORITE,PFI_SALARIE,PFI_LIBEMPLOIFOR,PFI_CODESTAGE,PFI_NATUREFORM,PFI_AUTRECOUT,PFI_ETATINSCFOR,PFI_MOTIFETATINSC'+
               ' ORDER BY PSI_LIBELLE,PFI_LIBELLE', True);
  //PT3 - Fin
  TFor := Tob.Create('Table',Nil,-1);
  TFor.LoadDetailDB('Table','','',Q,False);
  Ferme(Q);

  //PT3 - Début
  // On fusionne les deux listes
  If TForN <> Nil Then
  Begin
     For i:= TForN.Detail.Count-1 DownTo 0 Do
          TForN.Detail[i].ChangeParent(TFor,0,False);
     FreeAndNil(TForN);
     TFor.Detail.Sort('PSI_INTERIMAIRE;PFI_LIBELLE');
  End;
  //PT3 - Fin

  If TFor.Detail.Count = 0 then //PT3
     Grille.Enabled := False
  Else
     Grille.Enabled := True;
     
  TGrille := Tob.Create('Table',Nil,-1);
  TSomme := Tob.Create('sommes',Nil,-1);

  For i := 0 To TFor.Detail.Count - 1 Do
  Begin
     If TFor.Detail[i].GetValue('PSI_INTERIMAIRE') <> Resp Then    //PT3
     Begin
       //Insertion du compte-rendu du dernier responsable dans la grille
       If Resp <> '' Then //PT3
       Begin
          TS := Tob.create('FilleSomme',TSomme,-1);
          TS.AddChampSupValeur('SALARIE',Resp);
          TS.AddChampSupValeur('HEURES', NbHeures);
          TS.AddChampSupValeur('COUT',   Cout);
          TS.AddChampSupValeur('NBINSC', NbInsc);
       End;

       NbHeures := 0;
       NbInsc   := 0;
       Cout     := 0;
       Resp     := TFor.Detail[i].GetValue('PSI_INTERIMAIRE');
       
       TG := Tob.create('Fille',TGrille,-1);
       TG.AddChampSupValeur('COL1',TFor.Detail[i].GetValue('PSI_LIBELLE') + ' ' + TFor.Detail[i].GetValue('PSI_PRENOM'));
       TG.AddChampSupValeur('COL2','');
       TG.AddChampSupValeur('COL3','');
       TG.AddChampSupValeur('COL4','');
       TG.AddChampSupValeur('COL5','');
       TG.AddChampSupValeur('COL6','');
       TG.AddChampSupValeur('COL7','');
       TG.AddChampSupValeur('COL8','');
       TG.AddChampSupValeur('COL9','');
       TG.AddChampSupValeur('COL10','R');
       TG.AddChampSupValeur('COL11',Resp);
       TG.AddChampSupValeur('COL12','');
       TG.AddChampSupValeur('COL13','');
       TG.AddChampSupValeur('SALARIE',Resp);
     End;

     // Si PFI_LIBELLE non vide, il y a une inscription
     If TFor.Detail[i].GetValue('PFI_LIBELLE') <> '' Then //PT3
     Begin
            TG := Tob.create('Fille',TGrille,-1);
            If TFor.Detail[i].GetValue('PFI_SALARIE') <> '' then
               TG.AddChampSupValeur('COL1',TFor.Detail[i].GetValue('PFI_LIBELLE'))
            else
               TG.AddChampSupValeur('COL1','Non nominatif');
            TG.AddChampSupValeur('COL2',RechDom('PGLIBEMPLOI',TFor.Detail[i].GetValue('PFI_LIBEMPLOIFOR'),False));
            TG.AddChampSupValeur('COL3',RechDom('PGSTAGEFORM',TFor.Detail[i].GetValue('PFI_CODESTAGE'),False));
            TG.AddChampSupValeur('COL4',RechDom('PGNATUREFORM',TFor.Detail[i].GetValue('PFI_NATUREFORM'),False));
            TG.AddChampSupValeur('COL5',TFor.Detail[i].GetValue('PFI_NBINSC'));
            TG.AddChampSupValeur('COL6',TFor.Detail[i].GetValue('PFI_DUREESTAGE'));
            If TFor.Detail[i].GetValue('PFI_NATUREFORM') = '002' then
            begin
                 TG.AddChampSupValeur('COL7',TFor.Detail[i].GetValue('PFI_AUTRECOUT'));
                 Cout := Cout + TFor.Detail[i].GetValue('PFI_AUTRECOUT');
            end
            else TG.AddChampSupValeur('COL7',0);
            TG.AddChampSupValeur('COL8',RechDom('PGMOTIFINSCFORMATION',TFor.Detail[i].GetValue('PFI_MOTIFINSCFOR'),False));
            TG.AddChampSupValeur('COL9',RechDom('PGNIVPRIORITEFORM',TFor.Detail[i].GetValue('PFI_NIVPRIORITE'),False));
            TG.AddChampSupValeur('COL10','');
            TG.AddChampSupValeur('COL11','');
            TG.AddChampSupValeur('COL12',RechDom('PGETATVALIDATION',TFor.Detail[i].GetValue('PFI_ETATINSCFOR'),False));
            TG.AddChampSupValeur('COL13',RechDom('PGMOTIFETATINSC',TFor.Detail[i].GetValue('PFI_MOTIFETATINSC'),False));
            TG.AddChampSupValeur('SALARIE','');

            NbHeures  := NbHeures  + TFor.Detail[i].GetValue('PFI_DUREESTAGE');
            NbInsc    := NbInsc    + TFor.Detail[i].GetValue('PFI_NBINSC');
            //PT5 - Début
            //TotInsc   := TotInsc   + TFor.Detail[i].GetValue('PFI_NBINSC');
            //TotHeures := TotHeures + TFor.Detail[i].GetValue('PFI_DUREESTAGE');
            //TotHono   := TotHono   + Cout;
            //PT5 - Fin
       end;
  end;

  //Insertion du compte-rendu du dernier responsable dans la grille
  If Resp <> '' Then //PT3
  Begin
     TS := Tob.create('FilleSomme',TSomme,-1);
     TS.AddChampSupValeur('SALARIE',Resp);
     TS.AddChampSupValeur('HEURES', NbHeures);
     TS.AddChampSupValeur('COUT',   Cout);
     TS.AddChampSupValeur('NBINSC', NbInsc);
  End;

  TFor.Free;

  For i := 0 to TGrille.Detail.Count - 1 do
  begin
       if TGrille.Detail[i].GetValue('COL10') = 'R' then
       begin
            TS := TSomme.FindFirst(['SALARIE'],[TGrille.Detail[i].GetValue('SALARIE')],False);
            If TS <> Nil then
            begin
                 TGrille.Detail[i].PutValue('COL5',TS.GetValue('NBINSC'));
                 TGrille.Detail[i].PutValue('COL6',TS.GetValue('HEURES'));
                 TGrille.Detail[i].PutValue('COL7',TS.GetValue('COUT'));
                 
	            //PT5 - Début
	            TotInsc   := TotInsc   + TS.GetValue('NBINSC');
	            TotHeures := TotHeures + TS.GetValue('HEURES');
	            TotHono   := TotHono   + TS.GetValue('COUT');
	            //PT5 - Fin
            end;
       end;
  end;

  TSomme.Free;

  If TGrille.Detail.Count > 0 Then
     TGrille.PutGridDetail(THGrid(GetControl('GRILLE')),False,False,'COL1;COL2;COL3;COL4;COL5;COL6;COL7;COL8;COL9;COL10;COL11;COL12;COL13',False);
  TGrille.Free;

  Grille.CellValues[0,1] := 'Salarié';
  Grille.CellValues[1,1] := 'Emploi';
  Grille.CellValues[2,1] := 'Stage';
  Grille.CellValues[3,1] := 'Nature';
  For i := 2 to Grille.RowCount - 1 do
  begin
       If Grille.CellValues[9,i] = 'R' then Grille.RowHeights[i] := 24
       else Grille.RowHeights[i] := 15;
  end;
  Grille.DefaultColWidth := 100;

  //PT3 - Début
  Grille.Cells[0,0] :=  'Responsable';
  Grille.Cells[1,0] :=  '';
  Grille.Cells[2,0] :=  '';
  Grille.Cells[3,0] :=  '';
  Grille.Cells[4,0] :=  'Inscrits';
  Grille.Cells[5,0] :=  'Heures';
  Grille.Cells[6,0] :=  'Honoraires';
  Grille.Cells[7,0] :=  'Priorité';
  Grille.Cells[8,0] :=  'Prévision';
  Grille.Cells[9,0] :=  '';
  Grille.Cells[10,0]:=  '';
  Grille.Cells[11,0]:=  'Etat';
  Grille.Cells[12,0]:=  'Motif';

  Grille.ColFormats[2] := '';
  Grille.ColFormats[3] := '';
  Grille.ColFormats[4] := '# ##0';
  Grille.ColFormats[5] := '# ##0.00';
  Grille.ColFormats[6] := '# ##0';

  Grille.ColAligns[2] := taLeftJustify;
  Grille.ColAligns[3] := taLeftJustify;
  Grille.ColAligns[4] := taRightJustify;
  Grille.ColAligns[5] := taRightJustify;
  Grille.ColAligns[6] := taRightJustify;

  Grille.ColWidths[0] := 150;
  Grille.ColWidths[1] := 150;
  Grille.ColWidths[2] := 150;
  Grille.ColWidths[3] := 60;
  Grille.ColWidths[4] := 50;
  Grille.ColWidths[5] := 50;
  Grille.ColWidths[6] := 50;
  Grille.ColWidths[7] := 50;
  Grille.ColWidths[8] := 50;
  Grille.ColWidths[9] := -1;
  Grille.ColWidths[10]:= -1;
  Grille.ColWidths[11]:= 50;
  Grille.ColWidths[12]:= 70;
  //PT3 - Fin

  TFVierge(Ecran).HMTrad.ResizeGridColumns(Grille) ;
  SetControlText('TOTINSC',floatToStr(TotInsc));
  SetControlText('TOTHEURES',floatToStr(TotHeures));
  SetControlText('TOTHONO',floatToStr(Tothono));
end;
*)


procedure TOF_PGTESTFORMATION.GrilleDblClick(Sender : TObject);
var Salarie,Formation,LibEmploi,Cursus : String; //PT1
    i : Integer;
    Stconsult : String;
    Col : Integer; //PT5
begin
	 //PT5 - Début
	 If GetCheckBoxState('MULTINIVEAU') = CbChecked Then
	 	Col := 8
	 Else
	 	Col := 7;
	 //PT5 - Fin
     If Consultation then StConsult := ';ACTION=CONSULTATION';
     If GetControlText('CREGROUPEMENT') = 'FOR' then
     begin
          //PT3 - Début //PT5
          (*If GetControlText('MULTINIVEAU') = 'X' Then
          Begin
             PGIInfo('Veuillez décocher la visualisation des sous-niveaux, puis sélectionner la formation pour y inscrire des salariés','Information');
             Exit;
          End;*)
          //PT3 - Fin
          If Grille.CellValues[Col,Grille.Row] = 'F' then
          begin
               formation := Grille.CellValues[Col+1,Grille.Row];
               AglLanceFiche('PAY','EM_SAISIESTAGE','','','CWASINSCBUDGET;;'+formation+';'+Millesime+';;;'+StConsult+';'+Utilisateur);
               ChangeRegroupement(Nil);
          end
          else
          begin
               For i := Grille.Row - 1 downto 2 do
               begin
                    If Grille.CellValues[Col,i] = 'F' then
                    begin
                         formation := Grille.CellValues[Col+1,i];
                         AglLanceFiche('PAY','EM_SAISIESTAGE','','','CWASINSCBUDGET;;'+formation+';'+Millesime+';;;'+StConsult+';'+Utilisateur);
                         ChangeRegroupement(Nil);
                         exit;
                    end
                end;
          end;
     end
     //PT1 - Début
     Else If GetControlText('CREGROUPEMENT') = 'CUR' then
     begin
          //PT3 - Début //PT5
          (*If GetControlText('MULTINIVEAU') = 'X' Then
          Begin
             PGIInfo('Veuillez décocher la visualisation des sous-niveaux, puis sélectionner le cursus pour y inscrire des salariés','Information');
             Exit;
          End;*)
          //PT3 - Fin
          If Grille.CellValues[Col,Grille.Row] = 'C' then
          begin
               Cursus := Grille.CellValues[Col+1,Grille.Row];
               AglLanceFiche('PAY','EM_SAISIECURSUS','','','CURSUS;'+Cursus+';'+Millesime+';'+Utilisateur);
               ChangeRegroupement(Nil);
          end
          else
          begin
               For i := Grille.Row - 1 downto 2 do
               begin
                    If Grille.CellValues[Col,i] = 'C' then
                    begin
                         Cursus := Grille.CellValues[Col+1,i];
                         AglLanceFiche('PAY','EM_SAISIECURSUS','','','CURSUS;'+Cursus+';'+Millesime+';'+Utilisateur);
                         ChangeRegroupement(Nil);
                         Exit;
                    End
                End;
          End;
     end
     //PT1 - Fin
     else if GetControlText('CREGROUPEMENT') = 'LIBEMP' then
     begin
          //PT3 - Début //PT5
          (*If GetControlText('MULTINIVEAU') = 'X' Then
          Begin
             PGIInfo('Veuillez décocher la visualisation des sous-niveaux, puis sélectionner le salarié pour l''inscrire à une formation','Information');
             Exit;
          End;*)
          //PT3 - Fin
          If Grille.CellValues[Col,Grille.Row] = 'L' then
          begin
               LibEmploi := Grille.CellValues[Col+1,Grille.Row];
               AglLanceFiche('PAY','EM_SAISIESALNM','','','CWASINSCBUDGET;;;'+Millesime+';;'+LibEmploi+';'+StConsult+';'+Utilisateur);
               ChangeRegroupement(Nil);
          end
          else
          begin
               For i := Grille.Row - 1 downto 2 do
               begin
                    If Grille.CellValues[Col,i] = 'L' then
                    begin
                         LibEmploi := Grille.CellValues[Col+1,i];
                         AglLanceFiche('PAY','EM_SAISIESALNM','','','CWASINSCBUDGET;;;'+Millesime+';;'+LibEmploi+';'+StConsult+';'+Utilisateur);
                         ChangeRegroupement(Nil);
                         exit;
                    end
                end;
          end;
     end
     else
     begin
          //PT3 - Début //PT5
          (*If GetControlText('MULTINIVEAU') = 'X' Then
          Begin
             PGIInfo('Veuillez décocher la visualisation des sous-niveaux, puis sélectionner le salarié pour l''inscrire à une formation','Information');
             Exit;
          End;*)
          //PT3 - Fin
          If Grille.CellValues[Col,Grille.Row] = 'S' then
          begin
               Salarie := Grille.CellValues[Col+1,Grille.Row];
               if Salarie <> 'NULL' then  AglLanceFiche('PAY','EM_SAISIESAL','','','CWASINSCBUDGET;;;'+Millesime+';'+Salarie+';;'+StConsult+';'+Utilisateur)
               else AglLanceFiche('PAY','EM_SAISIESALNM','','','CWASINSCBUDGET;;;'+Millesime+';;;;'+Utilisateur);
               ChangeRegroupement(Nil);
          end
          else
          begin
               For i := Grille.Row - 1 downto 2 do
               begin
                    If Grille.CellValues[Col,i] = 'S' then
                    begin
                         Salarie := Grille.CellValues[Col+1,i];
                         if Salarie <> 'NULL' then  AglLanceFiche('PAY','EM_SAISIESAL','','','CWASINSCBUDGET;;;'+Millesime+';'+Salarie+';;'+StConsult+';'+Utilisateur)
                         else AglLanceFiche('PAY','EM_SAISIESALNM','','','CWASINSCBUDGET;;;'+Millesime+';;;'+StConsult+';'+Utilisateur);
                         ChangeRegroupement(Nil);
                         exit;
                    end
                end;
          end;
     end;
end;

procedure TOF_PGTESTFORMATION.BChercheClick(Sender : TObject);
begin
     ChangeRegroupement(Nil);
     SetControlEnabled('BCHERCHE',False);
end;

procedure TOF_PGTESTFORMATION.ChangeCritere(Sender : Tobject);
begin
     SetControlEnabled('BCHERCHE',True);
end;

procedure TOF_PGTESTFORMATION.AffichageParLibEmploi;
var Q : TQuery;
    TG,TGrille,TFor,TForN : Tob;
    i : Integer;
    LibEmploi,Where : String;
    NbHeures,Cout : Double;
    NbInsc : Integer;
    TSomme,TS : Tob;
    TotHeures,TotSal,TotInsc,Tothono : Double;
    NbCol, Col : Integer; //PT5
begin
  TotSal    := 0;
  TotInsc   := 0;
  TotHeures := 0;
  TotHono   := 0;
  NbInsc    := 0;
  Cout      := 0;
  NbHeures  := 0;

  For i := 0 to Grille.RowCount - 1 do Grille.Rows[i].Clear;
  Grille.RowCount := 3;
  //PT5 - Début
  If (GetCheckBoxState('MULTINIVEAU')=CbChecked) Then
    NbCol := 12
  Else
    NbCol := 11; //PT3
  Grille.ColCount := NbCol;
  //PT5 - Fin

  //PT3 - Début
  // Cas particulier des inscriptions non nominatives
  // ------------------------------------------------
  If GetControlText('PSA_SALARIE') = '' Then
  Begin
       Where := '';
       If GetControlText('PSA_LIBELLEEMPLOI') <> '' Then Where := Where + ' AND PFI_LIBEMPLOIFOR="'+GetControlText('PSA_LIBELLEEMPLOI')+'"';
       Where := Where + ' AND ' + AdaptByRespEmanager (TypeUtilisat,'PFI',Utilisateur,(GetCheckBoxState('MULTINIVEAU')=CbChecked)); //PT5

       Q := OpenSQL('SELECT DISTINCT(PFI_LIBEMPLOIFOR) AS PSA_LIBELLEEMPLOI,"NULL" AS PFI_SALARIE,'+
                    'PFI_LIBELLE,PFI_MOTIFINSCFOR,PFI_NIVPRIORITE,PFI_CODESTAGE,PFI_AUTRECOUT,PFI_NATUREFORM,PFI_RESPONSFOR,PFI_NBINSC,PFI_DUREESTAGE,PFI_ETATINSCFOR,PFI_MOTIFETATINSC '+ //PT5
                    'FROM INSCFORMATION '+
                    'WHERE PFI_MILLESIME="'+Millesime+'" AND PFI_CODESTAGE<>"--CURSUS--" AND PFI_SALARIE=""'+Where+
                    ' ORDER BY PFI_LIBELLE',True);

       TForN := Tob.Create('Table',Nil,-1);
       TForN.LoadDetailDB('Table','','',Q,False);
       Ferme(Q);
  end;
  //PT3 - Fin

  // Autres inscriptions
  // -------------------
  Where := '';
  //PT1 - Début
  //PT3 - Début
  If GetControlText('PSA_LIBELLEEMPLOI') <> '' Then Where := Where + ' AND PSA_LIBELLEEMPLOI="'+GetControlText('PSA_LIBELLEEMPLOI')+'"';
  If GetControlText('PSA_SALARIE') <> '' Then Where := Where + ' AND PSE_SALARIE="' + GetControlText('PSA_SALARIE')+'"';
  Where := Where + ' AND ' + AdaptByRespEmanager (TypeUtilisat,'PSE',Utilisateur,(GetCheckBoxState('MULTINIVEAU')=CbChecked)); //PT5

  Q := OpenSQL('SELECT DISTINCT(PSA_LIBELLEEMPLOI),'+
               'PFI_SALARIE,PFI_LIBELLE,PFI_MOTIFINSCFOR,PFI_NIVPRIORITE,PFI_CODESTAGE,PFI_AUTRECOUT,PFI_NATUREFORM,PFI_RESPONSFOR,PFI_NBINSC,PFI_DUREESTAGE,PFI_ETATINSCFOR,PFI_MOTIFETATINSC '+ //PT5
               'FROM DEPORTSAL LEFT JOIN INSCFORMATION ON (PSE_SALARIE=PFI_SALARIE AND PFI_MILLESIME="'+Millesime+'" AND PFI_CODESTAGE<>"--CURSUS--")'+
                             ' LEFT JOIN SALARIES ON (PSE_SALARIE=PSA_SALARIE) '+
               'WHERE (PSA_DATESORTIE IS NULL OR PSA_DATESORTIE>="'+UsdateTime(DateSortieSal)+'" OR PSA_DATESORTIE<="'+UsDateTime(IDate1900)+'")'+Where+
               ' ORDER BY PSA_LIBELLEEMPLOI,PFI_LIBELLE', True);
  //PT3 - Fin
  //PT1 - Fin

  TFor := Tob.Create('Table',Nil,-1);
  TFor.LoadDetailDB('Table','','',Q,False);
  Ferme(Q);

  //PT3 - Début
  // On fusionne les deux listes
  If TForN <> Nil Then
  Begin
     For i:= TForN.Detail.Count-1 DownTo 0 Do
          TForN.Detail[i].ChangeParent(TFor,0,False);
     FreeAndNil(TForN);
     TFor.Detail.Sort('PSA_LIBELLEEMPLOI');
  End;
  //PT3 - Fin

  // Si pas de données : Grille vide
  If TFor.Detail.Count = 0 then
     Grille.Enabled := False
  Else
     Grille.Enabled := True;

  TGrille := Tob.Create('Table',Nil,-1);
  TSomme  := Tob.Create('sommes',Nil,-1);

  For i := -0 to TFor.Detail.Count - 1 do
  begin
     If TFor.Detail[i].GetValue('PSA_LIBELLEEMPLOI') <> LibEmploi Then    //PT3
     Begin
       //Insertion du compte-rendu du dernier salarié dans la grille
       If LibEmploi <> '' Then //PT3
       Begin
          TS := Tob.create('FilleSomme',TSomme,-1);
          TS.AddChampSupValeur('LIBEMPLOI',LibEmploi);
          TS.AddChampSupValeur('HEURES',   NbHeures);
          TS.AddChampSupValeur('COUT',     Cout);
          TS.AddChampSupValeur('NBINSC',   NbInsc);
       End;
       
       NbHeures := 0;
       NbInsc   := 0;
       Cout     := 0;

       TG := Tob.create('Fille',TGrille,-1);
       LibEmploi := TFor.Detail[i].GetValue('PSA_LIBELLEEMPLOI');
       TG.AddChampSupValeur('COL1',RechDom('PGLIBEMPLOI',TFor.Detail[i].GetValue('PSA_LIBELLEEMPLOI'),False));
       TG.AddChampSupValeur('COL2','');
       TG.AddChampSupValeur('COL2BIS', ''); //PT5
       TG.AddChampSupValeur('COL3','');
       TG.AddChampSupValeur('COL4','');
       TG.AddChampSupValeur('COL5','');
       TG.AddChampSupValeur('COL6','');
       TG.AddChampSupValeur('COL7','');
       TG.AddChampSupValeur('COL8','L');
       TG.AddChampSupValeur('COL9',LibEmploi);
       TG.AddChampSupValeur('COL10','');
       TG.AddChampSupValeur('COL11','');
       TG.AddChampSupValeur('SALARIE',LibEmploi);
     End;

     // Si PFI_LIBELLE non vide, il y a une inscription
     If TFor.Detail[i].GetValue('PFI_LIBELLE') <> '' Then //PT3
     Begin
            TG := Tob.create('Fille',TGrille,-1);
            If TFor.Detail[i].GetValue('PFI_SALARIE') <> '' then
               TG.AddChampSupValeur('COL1',TFor.Detail[i].GetValue('PFI_LIBELLE'))
            else
               TG.AddChampSupValeur('COL1','Inscriptions non nominatives');
            TG.AddChampSupValeur('COL2',RechDom('PGSTAGEFORM',TFor.Detail[i].GetValue('PFI_CODESTAGE'),False));
            TG.AddChampSupValeur('COL2BIS', RechDom('PGINTERIMAIRES', TFor.Detail[i].GetValue('PFI_RESPONSFOR'),False)); //PT5
            TG.AddChampSupValeur('COL3',TFor.Detail[i].GetValue('PFI_NBINSC'));
            TG.AddChampSupValeur('COL4',TFor.Detail[i].GetValue('PFI_DUREESTAGE'));
            If TFor.Detail[i].GetValue('PFI_NATUREFORM') = '002' then
            begin
                 TG.AddChampSupValeur('COL5',TFor.Detail[i].GetValue('PFI_AUTRECOUT'));
                 Cout := Cout + TFor.Detail[i].GetValue('PFI_AUTRECOUT');
            end
            else TG.AddChampSupValeur('COL5',0);
            TG.AddChampSupValeur('COL6',RechDom('PGMOTIFINSCFORMATION',TFor.Detail[i].GetValue('PFI_MOTIFINSCFOR'),False));
            TG.AddChampSupValeur('COL7',RechDom('PGNIVPRIORITEFORM',TFor.Detail[i].GetValue('PFI_NIVPRIORITE'),False));
            TG.AddChampSupValeur('COL8','');
            TG.AddChampSupValeur('COL9','');
            TG.AddChampSupValeur('COL10',RechDom('PGETATVALIDATION',TFor.Detail[i].GetValue('PFI_ETATINSCFOR'),False));
            TG.AddChampSupValeur('COL11',RechDom('PGMOTIFETATINSC',TFor.Detail[i].GetValue('PFI_MOTIFETATINSC'),False));
            TG.AddChampSupValeur('SALARIE','');

            NbHeures  := NbHeures  + TFor.Detail[i].GetValue('PFI_DUREESTAGE');
            NbInsc    := NbInsc    + TFor.Detail[i].GetValue('PFI_NBINSC');
            //PT5 - Début
            //TotInsc   := TotInsc   + TFor.Detail[i].GetValue('PFI_NBINSC');
            //TotHeures := TotHeures + TFor.Detail[i].GetValue('PFI_DUREESTAGE');
            //TotHono   := TotHono   + Cout;
            //PT5 - Fin
       end;
  end;
  //Insertion du compte-rendu du dernier salarié dans la grille
  If LibEmploi <> '' Then //PT3
  Begin
     TS := Tob.create('FilleSomme',TSomme,-1);
     TS.AddChampSupValeur('LIBEMPLOI',LibEmploi);
     TS.AddChampSupValeur('HEURES',   NbHeures);
     TS.AddChampSupValeur('COUT',     Cout);
     TS.AddChampSupValeur('NBINSC',   NbInsc);
  End;
  TFor.Free;

  For i := 0 to TGrille.Detail.Count - 1 do
  begin
       if TGrille.Detail[i].GetValue('COL8') = 'L' then
       begin
            TS := TSomme.FindFirst(['LIBEMPLOI'],[TGrille.Detail[i].GetValue('SALARIE')],False);
            If TS <> Nil then
            begin
                 TotSal := TotSal + 1;
                 TGrille.Detail[i].PutValue('COL3',TS.GetValue('NBINSC'));
                 TGrille.Detail[i].PutValue('COL4',TS.GetValue('HEURES'));
                 TGrille.Detail[i].PutValue('COL5',TS.GetValue('COUT'));
                 
	            //PT5 - Début
	            TotInsc   := TotInsc   + TS.GetValue('NBINSC');
	            TotHeures := TotHeures + TS.GetValue('HEURES');
	            TotHono   := TotHono   + TS.GetValue('COUT');
	            //PT5 - Fin
            end;
       end;
  end;

  TSomme.Free;

  If TGrille.Detail.Count > 0 Then //PT3
  	If (GetCheckBoxState('MULTINIVEAU')=CbChecked) Then  //PT5
     TGrille.PutGridDetail(THGrid(GetControl('GRILLE')),False,False,'COL1;COL2;COL2BIS;COL3;COL4;COL5;COL6;COL7;COL8;COL9;COL10;COL11',False) 
    else
     TGrille.PutGridDetail(THGrid(GetControl('GRILLE')),False,False,'COL1;COL2;COL3;COL4;COL5;COL6;COL7;COL8;COL9;COL10;COL11',False);

  TGrille.Free;
  Grille.CellValues[0,1] := 'Salarié';
  Grille.CellValues[1,1] := 'Formation';
  
  //PT5 - Début
  If (GetCheckBoxState('MULTINIVEAU')=CbChecked) Then 
  	Col := 8
  else
  	Col := 7;
  //PT5 -Fin
  For i := 2 to Grille.RowCount - 1 do
  begin
       If Grille.CellValues[Col,i] = 'L' then Grille.RowHeights[i] := 24
       else Grille.RowHeights[i] := 15;
  end;

  //PT3 - Début
  Grille.Cells[0,0] :=  'Libellé emploi';
  Grille.Cells[1,0] :=  '';
  Col := 2;
  If NbCol = 12 Then Begin Grille.Cells[Col,0] := 'Responsable'; Col := Col + 1; End; //PT5
  Grille.Cells[Col,0] :=  'Inscrits';			Col := Col + 1;
  Grille.Cells[Col,0] :=  'Heures';             Col := Col + 1;
  Grille.Cells[Col,0] :=  'Honoraires';         Col := Col + 1;
  Grille.Cells[Col,0] :=  'Priorité';           Col := Col + 1;
  Grille.Cells[Col,0] :=  'Prévision';          Col := Col + 1;
  Grille.Cells[Col,0] :=  '';                   Col := Col + 1;
  Grille.Cells[Col,0] :=  '';                   Col := Col + 1;
  Grille.Cells[Col,0] :=  'Etat';               Col := Col + 1;
  Grille.Cells[Col,0]:=  'Motif';                
                                                  
  Grille.ColWidths[0] := 200;                      
  Grille.ColWidths[1] := 150;                       
  Col := 2;                                          
  If NbCol = 12 Then Begin Grille.ColWidths[2] := 100; Col := Col + 1; End; //PT5
  Grille.ColWidths[Col] := 50;					Col := Col + 1;
  Grille.ColWidths[Col] := 50;                  Col := Col + 1;
  Grille.ColWidths[Col] := 50;                  Col := Col + 1;
  Grille.ColWidths[Col] := 50;                  Col := Col + 1;
  Grille.ColWidths[Col] := 50;                  Col := Col + 1;
  Grille.ColWidths[Col] := -1;                  Col := Col + 1;
  Grille.ColWidths[Col] := -1;                  Col := Col + 1;
  Grille.ColWidths[Col] := 50;                  Col := Col + 1;
  Grille.ColWidths[Col] := 70;                             

  Col := 2;
  If NbCol = 12 Then Col := Col + 1; //PT5
  Grille.ColFormats[Col] := '# ##0';            Col := Col + 1;
  Grille.ColFormats[Col] := '# ##0.00';         Col := Col + 1;
  Grille.ColFormats[Col] := '# ##0';            Col := Col + 1;
  Grille.ColFormats[Col] := '';                 Col := Col + 1;
  Grille.ColFormats[Col] := '';

  Col := 2;
  If NbCol = 12 Then Begin Grille.ColAligns[Col] := taLeftJustify; Col := Col + 1; End; //PT5
  Grille.ColAligns[Col] := taRightJustify;		Col := Col + 1;
  Grille.ColAligns[Col] := taRightJustify;		Col := Col + 1;
  Grille.ColAligns[Col] := taRightJustify;		Col := Col + 1;
  Grille.ColAligns[Col] := taLeftJustify;		Col := Col + 1;
  Grille.ColAligns[Col] := taLeftJustify;
  //PT3 - Fin

  TFVierge(Ecran).HMTrad.ResizeGridColumns(Grille) ;
  SetControlText('TOTINSC',floatToStr(TotInsc));
  SetControlText('TOTHEURES',floatToStr(TotHeures));
  SetControlText('TOTHONO',floatToStr(Tothono));
end;

procedure TOF_PGTESTFORMATION.SalarieElipsisClick(Sender: TObject);
var
  StFrom, StWhere: string;
begin
  StWhere := '(PSA_DATESORTIE<="' + UsDateTime(IDate1900) + '" OR PSA_DATESORTIE IS NULL OR PSA_DATESORTIE>="' + UsDateTime(V_PGI.DateEntree) + '")';
  StFrom := 'SALARIES LEFT JOIN DEPORTSAL ON PSA_SALARIE=PSE_SALARIE';

  //PT1 - Début
{  If TypeUtilisat = 'R' then
  Begin
    if GetCheckBoxState('MULTINIVEAU')<>CbChecked then
      StWhere := StWhere + ' AND (PSE_RESPONSFOR="'+V_PGI.UserSalarie+'")'
    else
      StWhere := StWhere + ' AND (PSE_RESPONSFOR="'+V_PGI.UserSalarie+'"'+
      ' OR PSE_RESPONSFOR IN (SELECT PGS_RESPONSFOR FROM SERVICES where PGS_CODESERVICE IN '+
      '(SELECT PSO_CODESERVICE FROM SERVICES LEFT JOIN SERVICEORDRE ON PGS_CODESERVICE=PSO_SERVICESUP AND'+
      ' PGS_RESPONSFOR="'+V_PGI.UserSalarie+'")))';
  End
  else
  Begin
    if GetCheckBoxState('MULTINIVEAU')<>CbChecked then
      StWhere := StWhere + ' AND PSE_RESPONSFOR IN '+
      '(SELECT PGS_RESPONSFOR FROM SERVICES WHERE PGS_SECRETAIREFOR="'+V_PGI.UserSalarie+'")'
    else
      StWhere := StWhere + ' AND PSE_RESPONSFOR IN '+
      '(SELECT PGS_RESPONSFOR FROM SERVICES LEFT JOIN SERVICEORDRE ON PGS_CODESERVICE=PSO_CODESERVICE '+
      'WHERE (PGS_SECRETAIREFOR="'+V_PGI.UserSalarie+'" OR PSO_SERVICESUP IN (SELECT PGS_CODESERVICE FROM SERVICES WHERE PGS_SECRETAIREFOR="'+V_PGI.UserSalarie+'")))';
  End;}
  StWhere := StWhere + ' AND ' + AdaptByRespEmanager (TypeUtilisat,'PSE',Utilisateur,(GetCheckBoxState('MULTINIVEAU')=CbChecked));
  //PT1 - Fin
  LookupList(THEdit(Sender), 'Liste des salariés', StFrom, 'PSA_SALARIE', 'PSA_LIBELLE,PSA_PRENOM', StWhere, 'PSA_SALARIE', TRUE, -1);
end;

procedure TOF_PGTESTFORMATION.RespElipsisClick(Sender: TObject);
var StFrom, StWhere: string;
begin
  THEdit(Sender).Text := '';
  StFrom := 'INTERIMAIRES';
  //PT1 - Début
   If TypeUtilisat = 'R' then
    StWhere := StWhere + ' AND (PSI_INTERIMAIRE="'+V_PGI.UserSalarie+'" '+
    'OR PSI_INTERIMAIRE IN (SELECT PGS_RESPONSFOR FROM SERVICES where PGS_CODESERVICE IN '+
    '(SELECT PSO_CODESERVICE FROM SERVICES LEFT JOIN SERVICEORDRE ON PGS_CODESERVICE=PSO_SERVICESUP AND '+
    'PGS_RESPONSFOR="'+V_PGI.UserSalarie+'")))'
   Else
    StWhere := StWhere + ' AND (PSI_INTERIMAIRE IN (SELECT PGS_RESPONSFOR FROM SERVICES WHERE PGS_SECRETAIREFOR="'+V_PGI.UserSalarie+'" OR PGS_CODESERVICE IN '+
    '(SELECT PSO_CODESERVICE FROM SERVICES LEFT JOIN SERVICEORDRE ON PGS_CODESERVICE=PSO_SERVICESUP AND '+
    'PGS_SECRETAIREFOR="'+V_PGI.UserSalarie+'")))';
  //PT1 - Fin
  LookupList(THEdit(Sender), 'Liste des responsables', StFrom, 'PSI_INTERIMAIRE', 'PSI_LIBELLE,PSI_PRENOM', StWhere, 'PSI_INTERIMAIRE', TRUE, -1);

  If GetControlText('RESPONSFOR') <> '' Then
     Utilisateur := GetControlText('RESPONSFOR')
  Else
     Utilisateur := V_PGI.UserSalarie;
  If ExisteSQL('SELECT PSE_SALARIE FROM DEPORTSAL WHERE PSE_RESPONSFOR="'+Utilisateur+'"') then TypeUtilisat := 'R'
  else TypeUtilisat := 'S';


  SetControlProperty('PSA_LIBELLEEMPLOI','Plus', ' AND CC_CODE IN (SELECT PSA_LIBELLEEMPLOI FROM SALARIES LEFT JOIN DEPORTSAL ON PSA_SALARIE=PSE_SALARIE '+
  'WHERE (PSA_DATESORTIE IS NULL OR PSA_DATESORTIE>="'+UsdateTime(V_PGI.DateEntree)+'" OR PSA_DATESORTIE<="'+UsDateTime(IDate1900)+'") AND '+
  AdaptByRespEmanager (TypeUtilisat,'PSE',Utilisateur,False)+')');
end;

//PT4 - Début
procedure TOF_PGTESTFORMATION.ExportGrille(Sender : TObject);
var SD: TSaveDialog;
    NomFic : String;
    Salarie,Where,SQL : String;
    TFor,TGrille,TG : Tob;
    i : Integer;
begin
     // Nom du fichier de sauvegarde
     SD:=TSaveDialog.Create(nil);
     SD.Filter:='Fichier Excel (*.xls)|*.xls';

     // Ajout de l'extension
     If SD.Execute Then
     Begin
          If Pos('.xls',Sd.FileName) = 0 Then
               NomFic := Sd.FileName+'.xls'
          Else
               NomFic := Sd.FileName;
     End;
     SD.Destroy;
     If NomFic = '' Then Exit;

     Where := '';

     (*If GetCheckBoxState('MULTINIVEAU') = CbChecked then		//PT5
     begin
          // Critères de sélection
          If GetControlText('PSA_SALARIE') <> '' then Where := Where + ' AND PFI_SALARIE="'+GetControlText('PSA_SALARIE')+'"';
          If GetControlText('PSA_LIBELLEEMPLOI') <> '' then Where := Where + ' AND PFI_LIBEMPLOIFOR="'+GetControlText('PSA_LIBELLEEMPLOI')+'"';
          Where := Where + ' AND ' + AdaptByRespEmanager (TypeUtilisat,'PFI',Utilisateur,True);

          // Chargement des informations
          TFor := TOB.Create('LesInfos', Nil, -1);
          TFor.LoadDetailFromSQL('SELECT PSI_INTERIMAIRE,PSI_LIBELLE,PSI_PRENOM,'+
               'PFI_SALARIE,PFI_LIBELLE,PFI_LIBEMPLOIFOR,PFI_MOTIFINSCFOR,PFI_NIVPRIORITE,PFI_CODESTAGE,PFI_AUTRECOUT,PFI_NATUREFORM,PFI_NBINSC,PFI_DUREESTAGE,PFI_ETATINSCFOR,PFI_MOTIFETATINSC '+
               'FROM INTERIMAIRES '+
               'LEFT JOIN INSCFORMATION ON (PFI_RESPONSFOR=PSI_INTERIMAIRE AND PFI_MILLESIME="'+Millesime+'" AND PFI_CODESTAGE<>"--CURSUS--") '+ //PT1
               'WHERE '+Where+
               ' ORDER BY PSI_LIBELLE,PFI_LIBELLE');

          // Formatage des informations
          TGrille := Tob.Create('Table',Nil,-1);
          For i := 0 to TFor.Detail.Count - 1 Do
          Begin
               TG := Tob.create('Fille',TGrille,-1);
               TG.AddChampSupValeur('RESPONSABLE',TFor.Detail[i].GetValue('PSI_LIBELLE') + ' '+TFor.Detail[i].GetValue('PSI_PRENOM'));
               Salarie := TFor.Detail[i].GetValue('PFI_SALARIE');
               If Salarie = '' then
                    TG.AddChampSupValeur('SALARIE',    'Inscriptions non nominatives')
               else
                    TG.AddChampSupValeur('SALARIE',    TFor.Detail[i].GetValue('PFI_LIBELLE'));
               TG.AddChampSupValeur('EMPLOI',     RechDom('PGLIBEMPLOI',TFor.Detail[i].GetValue('PFI_LIBEMPLOIFOR'),False));
               TG.AddChampSupValeur('STAGE',           RechDom('PGSTAGEFORM',TFor.Detail[i].GetValue('PFI_CODESTAGE'),False));
               TG.AddChampSupValeur('NATURE',     RechDom('PGNATUREFORM',TFor.Detail[i].GetValue('PFI_NATUREFORM'),False));
               TG.AddChampSupValeur('NBINSC',     TFor.Detail[i].GetValue('PFI_NBINSC'));
               TG.AddChampSupValeur('DUREE',      TFor.Detail[i].GetValue('PFI_DUREESTAGE'));
               If TFor.Detail[i].GetValue('PFI_NATUREFORM') = '002' Then
                    TG.AddChampSupValeur('HONORAIRES',TFor.Detail[i].GetValue('PFI_AUTRECOUT'))
               else
                    TG.AddChampSupValeur('HONORAIRES',0);
               TG.AddChampSupValeur('PRIORITE',   RechDom('PGMOTIFINSCFORMATION',TFor.Detail[i].GetValue('PFI_MOTIFINSCFOR'),False));
               TG.AddChampSupValeur('PREVISION',  RechDom('PGNIVPRIORITEFORM',TFor.Detail[i].GetValue('PFI_NIVPRIORITE'),False));
               TG.AddChampSupValeur('ETAT',       RechDom('PGETATVALIDATION',TFor.Detail[i].GetValue('PFI_ETATINSCFOR'),False));
               TG.AddChampSupValeur('MOTIF',      RechDom('PGMOTIFETATINSC',TFor.Detail[i].GetValue('PFI_MOTIFETATINSC'),False));
         end;
         FreeAndNil(TFor);

         TGrille.SaveToExcelFile(NomFic);
         FreeAndNil(TGrille);
  end
     else*) 
     If GetControlText('CREGROUPEMENT') = 'FOR' then
     begin
          // Critères de sélection
          If GetControlText('FAVORIS') = 'X' Then
               Where := ' AND PST_CODESTAGE IN (SELECT PCC_CODESTAGE FROM CURSUSSTAGE WHERE PCC_CURSUS="---" AND PCC_LIBELLE="'+V_PGI.UserSalarie+'") '
          Else
               Where := ' AND PST_CODESTAGE NOT IN (SELECT PCC_CODESTAGE FROM CURSUSSTAGE WHERE PCC_CURSUS="---" AND PCC_LIBELLE="'+V_PGI.UserSalarie+'") ';
          SQL := '';
          If GetControlText('PSA_SALARIE') <> '' Then SQL := SQL + ' AND PFI_SALARIE="'+GetControlText('PSA_SALARIE')+'"';
          If GetControlText('PSA_LIBELLEEMPLOI') <> '' Then SQL := SQL + ' AND PFI_LIBEMPLOIFOR="'+GetControlText('PSA_LIBELLEEMPLOI')+'"';
          SQL := SQL + ' AND ' + AdaptByRespEmanager (TypeUtilisat,'PFI',Utilisateur,GetCheckBoxState('MULTINIVEAU') = CbChecked);  //PT5

          // Chargement des informations
          TFor := TOB.Create('LesInfos', Nil, -1);
          TFor.LoadDetailFromSQL('SELECT PST_LIBELLE,PST_LIBELLE1,PST_NATUREFORM,'+
               'PFI_LIBELLE,PFI_LIBEMPLOIFOR,PFI_RESPONSFOR,'+ //PT5
               'PFI_NBINSC,PFI_DUREESTAGE,PFI_NATUREFORM,PFI_AUTRECOUT,PFI_MOTIFINSCFOR,PFI_NIVPRIORITE,PFI_ETATINSCFOR,PFI_MOTIFETATINSC '+
               'FROM STAGE INNER JOIN INSCFORMATION ON PST_CODESTAGE=PFI_CODESTAGE '+
               ' AND PFI_MILLESIME="'+Millesime+'" '+'AND '+
               SQL+
               ' WHERE '+
               'PST_ACTIF="X" AND PST_MILLESIME="0000" AND PST_FORMATION3="'+CYCLE_EN_COURS_EM+'"'+Where+
               ' ORDER BY PST_CODESTAGE,PFI_LIBELLE');

          // Formatage des informations
          TGrille := Tob.Create('Table',Nil,-1);
          For i := 0 to TFor.Detail.Count - 1 Do
          Begin
               TG := Tob.create('Fille',TGrille,-1);
               If GetControlText('RESPONSFOR') <> '' Then TG.AddChampSupValeur('RESPONSABLE',RechDom('PGINTERIMAIRES',Utilisateur,False));
               TG.AddChampSupValeur('STAGE',      TFor.Detail[i].GetValue('PST_LIBELLE') + ' ' + TFor.Detail[i].GetValue('PST_LIBELLE1'));
               TG.AddChampSupValeur('NATURE',     RechDom('PGNATUREFORM',TFor.Detail[i].GetValue('PFI_NATUREFORM'),False));
               TG.AddChampSupValeur('SALARIE',    TFor.Detail[i].GetValue('PFI_LIBELLE'));
               TG.AddChampSupValeur('EMPLOI',     RechDom('PGLIBEMPLOI',TFor.Detail[i].GetValue('PFI_LIBEMPLOIFOR'),False));
               TG.AddChampSupValeur('RESPONSABLE',RechDom('PGINTERIMAIRES',TFor.Detail[i].GetValue('PFI_RESPONSFOR'),False)); //PT5
               TG.AddChampSupValeur('NBINSC',     TFor.Detail[i].GetValue('PFI_NBINSC'));
               TG.AddChampSupValeur('DUREE',      TFor.Detail[i].GetValue('PFI_DUREESTAGE'));
               If TFor.Detail[i].GetValue('PFI_NATUREFORM') = '002' Then
                    TG.AddChampSupValeur('HONORAIRES',TFor.Detail[i].GetValue('PFI_AUTRECOUT'))
               else
                    TG.AddChampSupValeur('HONORAIRES',0);
               TG.AddChampSupValeur('PRIORITE',   RechDom('PGMOTIFINSCFORMATION',TFor.Detail[i].GetValue('PFI_MOTIFINSCFOR'),False));
               TG.AddChampSupValeur('PREVISION',  RechDom('PGNIVPRIORITEFORM',TFor.Detail[i].GetValue('PFI_NIVPRIORITE'),False));
               TG.AddChampSupValeur('ETAT',       RechDom('PGETATVALIDATION',TFor.Detail[i].GetValue('PFI_ETATINSCFOR'),False));
               TG.AddChampSupValeur('MOTIF',      RechDom('PGMOTIFETATINSC',TFor.Detail[i].GetValue('PFI_MOTIFETATINSC'),False));
         end;
         FreeAndNil(TFor);

         TGrille.SaveToExcelFile(NomFic);
         FreeAndNil(TGrille);
     end
     else
     If GetControlText('CREGROUPEMENT') = 'LIBEMP' then
     begin
          // Critères de sélection
          If GetControlText('PSA_LIBELLEEMPLOI') <> '' Then Where := Where + ' AND PSA_LIBELLEEMPLOI="'+GetControlText('PSA_LIBELLEEMPLOI')+'"';
          If GetControlText('PSA_SALARIE') <> '' Then Where := Where + ' AND PSE_SALARIE="' + GetControlText('PSA_SALARIE')+'"';
          Where := Where + ' AND ' + AdaptByRespEmanager (TypeUtilisat,'PFI',Utilisateur,GetCheckBoxState('MULTINIVEAU') = CbChecked); //PT5

          // Chargement des informations
          TFor := TOB.Create('LesInfos', Nil, -1);
          TFor.LoadDetailFromSQL('SELECT DISTINCT(PFI_LIBEMPLOIFOR),'+
               'PFI_LIBELLE,PFI_MOTIFINSCFOR,PFI_NIVPRIORITE,PFI_CODESTAGE,PFI_AUTRECOUT,PFI_NATUREFORM,PFI_RESPONSFOR,PFI_NBINSC,PFI_DUREESTAGE,PFI_ETATINSCFOR,PFI_MOTIFETATINSC '+ //PT5
               'FROM INSCFORMATION LEFT JOIN SALARIES ON (PFI_SALARIE=PSA_SALARIE) '+
               'WHERE (PSA_DATESORTIE IS NULL OR PSA_DATESORTIE>="'+UsdateTime(DateSortieSal)+'" OR PSA_DATESORTIE<="'+UsDateTime(IDate1900)+'")'+Where+
               ' AND PFI_MILLESIME="'+Millesime+'" AND PFI_CODESTAGE<>"--CURSUS--"'+
               ' ORDER BY PFI_LIBEMPLOIFOR', True);

          // Formatage des informations
          TGrille := Tob.Create('Table',Nil,-1);
          For i := 0 to TFor.Detail.Count - 1 Do
          Begin
               TG := Tob.create('Fille',TGrille,-1);
               //If GetControlText('RESPONSFOR') <> '' Then TG.AddChampSupValeur('RESPONSABLE',RechDom('PGINTERIMAIRES',Utilisateur,False)); //PT5
               TG.AddChampSupValeur('EMPLOI',     RechDom('PGLIBEMPLOI',TFor.Detail[i].GetValue('PFI_LIBEMPLOIFOR'),False));
               Salarie := TFor.Detail[i].GetValue('PFI_SALARIE');
               If Salarie = '' then
                    TG.AddChampSupValeur('SALARIE',    'Inscriptions non nominatives')
               else
                    TG.AddChampSupValeur('SALARIE',    TFor.Detail[i].GetValue('PFI_LIBELLE'));
               TG.AddChampSupValeur('STAGE',           RechDom('PGSTAGEFORM',TFor.Detail[i].GetValue('PFI_CODESTAGE'),False));
               TG.AddChampSupValeur('RESPONSABLE',     RechDom('PGINTERIMAIRES',TFor.Detail[i].GetValue('PFI_RESPONSFOR'),False)); //PT5
               TG.AddChampSupValeur('NBINSC',          TFor.Detail[i].GetValue('PFI_NBINSC'));
               TG.AddChampSupValeur('DUREE',           TFor.Detail[i].GetValue('PFI_DUREESTAGE'));
               If TFor.Detail[i].GetValue('PFI_NATUREFORM') = '002' Then
                    TG.AddChampSupValeur('HONORAIRES', TFor.Detail[i].GetValue('PFI_AUTRECOUT'))
               else
                    TG.AddChampSupValeur('HONORAIRES',0);
               TG.AddChampSupValeur('PRIORITE',        RechDom('PGMOTIFINSCFORMATION',TFor.Detail[i].GetValue('PFI_MOTIFINSCFOR'),False));
               TG.AddChampSupValeur('PREVISION',       RechDom('PGNIVPRIORITEFORM',TFor.Detail[i].GetValue('PFI_NIVPRIORITE'),False));
               TG.AddChampSupValeur('ETAT',            RechDom('PGETATVALIDATION',TFor.Detail[i].GetValue('PFI_ETATINSCFOR'),False));
               TG.AddChampSupValeur('MOTIF',           RechDom('PGMOTIFETATINSC',TFor.Detail[i].GetValue('PFI_MOTIFETATINSC'),False));
         end;
         FreeAndNil(TFor);

         TGrille.SaveToExcelFile(NomFic);
         FreeAndNil(TGrille);
     end
     Else If GetControlText('CREGROUPEMENT') = 'CUR' Then
     Begin
          // Critères de sélection
          If GetControlText('FAVORIS') = 'X' Then
               Where := ' AND PCU_CURSUS IN (SELECT PCC_CODESTAGE FROM CURSUSSTAGE WHERE PCC_CURSUS="-C-" AND PCC_LIBELLE="'+V_PGI.UserSalarie+'") '
          Else
               Where := ' AND PCU_CURSUS NOT IN (SELECT PCC_CODESTAGE FROM CURSUSSTAGE WHERE PCC_CURSUS="-C-" AND PCC_LIBELLE="'+V_PGI.UserSalarie+'") ';
          SQL := '';
          If GetControlText('PSA_SALARIE') <> '' Then SQL := SQL + ' AND PFI_SALARIE="'+GetControlText('PSA_SALARIE')+'"';
          If GetControlText('PSA_LIBELLEEMPLOI') <> '' Then SQL := SQL + ' AND PFI_LIBEMPLOIFOR="'+GetControlText('PSA_LIBELLEEMPLOI')+'"';
          SQL := SQL + ' AND ' + AdaptByRespEmanager (TypeUtilisat,'PFI',Utilisateur,GetCheckBoxState('MULTINIVEAU') = CbChecked); //PT5

          // Chargement des informations
          TFor := TOB.Create('LesInfos', Nil, -1);
          TFor.LoadDetailFromSQL('SELECT PCU_LIBELLE,PFI_LIBELLE,PFI_LIBEMPLOIFOR,PFI_RESPONSFOR,'+ //PT5
               'PFI_NBINSC,PFI_DUREESTAGE,PFI_NATUREFORM,PFI_AUTRECOUT,PFI_MOTIFINSCFOR,PFI_NIVPRIORITE,PFI_ETATINSCFOR,PFI_MOTIFETATINSC '+
               'FROM CURSUS INNER JOIN INSCFORMATION ON PCU_CURSUS=PFI_CURSUS AND PFI_MILLESIME="'+Millesime+'" AND PFI_CODESTAGE="--CURSUS--"'+
               SQL+
               ' WHERE PCU_RANGCURSUS=0 AND PCU_INCLUSCAT="X" '+Where+
               'ORDER BY PCU_CURSUS,PFI_LIBELLE');

          // Formatage des informations
          TGrille := Tob.Create('Table',Nil,-1);
          For i := 0 to TFor.Detail.Count - 1 Do
          Begin
               TG := Tob.create('Fille',TGrille,-1);
               If GetControlText('RESPONSFOR') <> '' Then TG.AddChampSupValeur('RESPONSABLE',RechDom('PGINTERIMAIRES',Utilisateur,False));
               TG.AddChampSupValeur('CURSUS',     TFor.Detail[i].GetValue('PCU_LIBELLE'));
               TG.AddChampSupValeur('SALARIE',    TFor.Detail[i].GetValue('PFI_LIBELLE'));
               TG.AddChampSupValeur('EMPLOI',     RechDom('PGLIBEMPLOI',TFor.Detail[i].GetValue('PFI_LIBEMPLOIFOR'),False));
               TG.AddChampSupValeur('RESPONSABLE',RechDom('PGINTERIMAIRES',TFor.Detail[i].GetValue('PFI_RESPONSFOR'),False)); //PT5
               TG.AddChampSupValeur('NBINSC',     TFor.Detail[i].GetValue('PFI_NBINSC'));
               TG.AddChampSupValeur('DUREE',      TFor.Detail[i].GetValue('PFI_DUREESTAGE'));
               If TFor.Detail[i].GetValue('PFI_NATUREFORM') = '002' Then
                    TG.AddChampSupValeur('HONORAIRES',TFor.Detail[i].GetValue('PFI_AUTRECOUT'))
               else
                    TG.AddChampSupValeur('HONORAIRES',0);
               TG.AddChampSupValeur('PRIORITE',   RechDom('PGMOTIFINSCFORMATION',TFor.Detail[i].GetValue('PFI_MOTIFINSCFOR'),False));
               TG.AddChampSupValeur('PREVISION',  RechDom('PGNIVPRIORITEFORM',TFor.Detail[i].GetValue('PFI_NIVPRIORITE'),False));
               TG.AddChampSupValeur('ETAT',       RechDom('PGETATVALIDATION',TFor.Detail[i].GetValue('PFI_ETATINSCFOR'),False));
               TG.AddChampSupValeur('MOTIF',      RechDom('PGMOTIFETATINSC',TFor.Detail[i].GetValue('PFI_MOTIFETATINSC'),False));
         end;
         FreeAndNil(TFor);

         TGrille.SaveToExcelFile(NomFic);
         FreeAndNil(TGrille);
     End
     else  // Salariés
     begin
          // Critères de sélection
          If GetControlText('PSA_LIBELLEEMPLOI') <> '' Then Where := Where + ' AND PSA_LIBELLEEMPLOI="'+GetControlText('PSA_LIBELLEEMPLOI')+'"';
          If GetControlText('PSA_SALARIE') <> '' Then Where := Where + ' AND PSE_SALARIE="' + GetControlText('PSA_SALARIE')+'"';
          Where := Where + ' AND ' + AdaptByRespEmanager (TypeUtilisat,'PFI',Utilisateur,GetCheckBoxState('MULTINIVEAU') = CbChecked); //PT5

          // Chargement des informations
          TFor := TOB.Create('LesInfos', Nil, -1);
          TFor.LoadDetailFromSQL('SELECT PFI_SALARIE,PFI_LIBELLE,PFI_LIBEMPLOIFOR,PFI_MOTIFINSCFOR,PFI_NIVPRIORITE,PFI_CODESTAGE,PFI_AUTRECOUT,PFI_NATUREFORM,'+
               'PFI_RESPONSFOR,PFI_NBINSC,PFI_DUREESTAGE,PFI_ETATINSCFOR,PFI_MOTIFETATINSC '+ //PT5
               'FROM INSCFORMATION LEFT JOIN SALARIES ON (PFI_SALARIE=PSA_SALARIE) '+
               'WHERE (PSA_DATESORTIE IS NULL OR PSA_DATESORTIE>="'+UsdateTime(DateSortieSal)+'" OR PSA_DATESORTIE<="'+UsDateTime(IDate1900)+'")'+Where+
               ' AND PFI_MILLESIME="'+Millesime+'" AND PFI_CODESTAGE<>"--CURSUS--"'+
               ' ORDER BY PSA_LIBELLE');

          // Formatage des informations
          TGrille := Tob.Create('Table',Nil,-1);
          For i := 0 to TFor.Detail.Count - 1 Do
          begin
               TG := Tob.create('Fille',TGrille,-1);
               Salarie := TFor.Detail[i].GetValue('PFI_SALARIE');
               If Salarie = '' then
                    TG.AddChampSupValeur('SALARIE',    'Inscriptions non nominatives')
               else
                    TG.AddChampSupValeur('SALARIE',    TFor.Detail[i].GetValue('PFI_LIBELLE'));
               TG.AddChampSupValeur('EMPLOI',          RechDom('PGLIBEMPLOI',TFor.Detail[i].GetValue('PFI_LIBEMPLOIFOR'),False));
               TG.AddChampSupValeur('STAGE',           RechDom('PGSTAGEFORM',TFor.Detail[i].GetValue('PFI_CODESTAGE'),False));
               TG.AddChampSupValeur('NATURE',          RechDom('PGNATUREFORM',TFor.Detail[i].GetValue('PFI_NATUREFORM'),False));
               TG.AddChampSupValeur('RESPONSABLE',     RechDom('PGINTERIMAIRES',TFor.Detail[i].GetValue('PFI_RESPONSFOR'),False)); //PT5
               TG.AddChampSupValeur('NBINSC',          TFor.Detail[i].GetValue('PFI_NBINSC'));
               TG.AddChampSupValeur('DUREE',           TFor.Detail[i].GetValue('PFI_DUREESTAGE'));
               If TFor.Detail[i].GetValue('PFI_NATUREFORM') = '002' then
                    TG.AddChampSupValeur('HONORAIRES',TFor.Detail[i].GetValue('PFI_AUTRECOUT'))
               else
                    TG.AddChampSupValeur('HONORAIRES',0);
               TG.AddChampSupValeur('PRIORITE',        RechDom('PGMOTIFINSCFORMATION',TFor.Detail[i].GetValue('PFI_MOTIFINSCFOR'),False));
               TG.AddChampSupValeur('PREVISION',       RechDom('PGNIVPRIORITEFORM',TFor.Detail[i].GetValue('PFI_NIVPRIORITE'),False));
               TG.AddChampSupValeur('ETAT',            RechDom('PGETATVALIDATION',TFor.Detail[i].GetValue('PFI_ETATINSCFOR'),False));
               TG.AddChampSupValeur('MOTIF',           RechDom('PGMOTIFETATINSC',TFor.Detail[i].GetValue('PFI_MOTIFETATINSC'),False));
          end;
          FreeAndNil(TFor);

          TGrille.SaveToExcelFile(NomFic);
          FreeAndNil(TGrille);
     end;
end;
//PT4 - Fin

Initialization
  registerclasses ( [ TOF_PGTESTFORMATION ] ) ;
end.


