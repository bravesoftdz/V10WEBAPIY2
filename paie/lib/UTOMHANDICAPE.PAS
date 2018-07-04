{***********UNITE*************************************************
Auteur  ...... : PH
Créé le ...... : 28/05/2001
Modifié le ... : 30/08/2001
Description .. : Source TOM gestion des handicapés
Mots clefs ... : PAIE;HANDICAPE
*****************************************************************
PT1 : 15/11/2001 : JL 562: Procedure ExitEdit pour afficher le code salarié complet
PT2 : 15/11/2001 : JL 562: Gestion du nouveau champ : date de fins de droits COTOREP
PT3 : 05/11/2002 : JL 591 : Correction affichage date au lieu de l'heure(propriété editMask
PT4 : 17/12/2002 : PH V591 Toutes les dates sont initialisées à idate1900 ou 2099 au lieu de null
PT5 : 20/05/2003 : JL V_42 Le format de la date déplacé dans la fiche mise en commentaire de PT3
PT6 : 15/11/2006 : JL V_70 FQ 13602 Pouvoir saisir les travailleurs handicapés qui sont à la fois reconnus par la COTOREP et par la Sécu
PT7 : 27/04/2007 : JL V_720 Gestion accès depuis ribbonbar salarié
}
Unit UTOMHANDICAPE ;

Interface

Uses  Controls, Classes, sysutils,
{$IFNDEF EAGLCLIENT}
     db, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}DBCtrls,HDB,
{$ELSE}
     eFiche,UtileAGL,StdCtrls,
{$ENDIF}
     HCtrls, HEnt1,  UTOM, UTob,EntPaie,PgOutils2;

Type
  TOM_HANDICAPE = Class (TOM)
    procedure OnNewRecord                ; override ;
    procedure OnLoadRecord               ; override ;
    procedure OnArgument ( S: String )   ; override ;
    private
    InitSalarie : String;
    procedure Handiclick(Sender : TObject);
    procedure COTOREPClick(Sender : TObject);
    procedure Autreclick(Sender : TObject);
    procedure affichenom(Sender : TObject);
    procedure ExitEdit(Sender : TObject);  //PT- 1
    end ;

Implementation


procedure TOM_HANDICAPE.OnNewRecord ;
begin
  Inherited ;
  If InitSalarie <> '' then          //PT7
  begin
    SetControlEnabled('PGH_SALARIE',False);
    SetField('PGH_SALARIE',InitSalarie);
  end;
end;

procedure TOM_HANDICAPE.OnLoadRecord ;
var
   QSAL : TQuery;
   nomsalarie,salarie : string;
begin
  Inherited ;
        if GetControlText('PGH_ACCTRAV') = 'X' then  SetControlEnabled('PGH_TXINV',True);
        salarie := GetField('PGH_SALARIE');
        QSAL := OpenSQL('SELECT PSA_LIBELLE FROM SALARIES WHERE PSA_SALARIE="'+salarie+'"',TRUE);
        nomsalarie := '';
        If Not QSAL.eof then nomsalarie := QSAL.FindField('PSA_LIBELLE').AsString;    // // PortageCWAS
        Ferme(QSAL);
        Ecran.caption := 'Saisie travailleurs handicapés de '+nomsalarie;
end ;

procedure TOM_HANDICAPE.OnArgument ( S: String ) ;
var
  Action : String;
{$IFNDEF EAGLCLIENT}
   Check1,Check2 : TDBCheckBox ;
   sal : THDBEdit;
   Check3,Check4,Check5,Check6: TDBCheckBox ;
{$ELSE}
   Check1,Check2 : TCheckBox ;
   sal : THEdit;
   Check3,Check4,Check5,Check6: TCheckBox ;
{$ENDIF}
begin
     Inherited ;
     Action := ReadTokenPipe(S,';');
     InitSalarie := ReadTokenPipe(S,';');   //PT7
{$IFNDEF EAGLCLIENT}
        sal := THDBEdit(GetControl('PGH_SALARIE')) ;
        Check1 := TDBCheckBox(GetControl('PGH_HANDICAP'));
        Check2 := TDBCheckBox(GetControl('PGH_COTOREP'));
{$ELSE}
        sal:=THEdit(GetControl('PGH_SALARIE')) ;
        Check1 := TCheckBox(GetControl('PGH_HANDICAP'));
        Check2 := TCheckBox(GetControl('PGH_COTOREP'));
{$ENDIF}

        if check1<>nil then Check1.OnClick:=Handiclick;
        if sal<>nil then
        begin
                sal.Onchange:=affichenom;
                sal.OnExit:=ExitEdit;   //PT- 1
        end;
        if check2<>nil then Check2.onclick:=COTOREPClick;
//SetControlProperty('PGH_DATEFINCOT','EditMask','!99/99/0000;1;_');    //PT- 3
             //Suivant les changements des CheckBox les différentes procédures sont appelés
{$IFNDEF EAGLCLIENT}
        check3 := TDBCheckBox(GetControl('PGH_ACCTRAV'));
        check4 := TDBCheckBox(GetControl('PGH_INVPEN'));
        check5 := TDBCheckBox(GetControl('PGH_MUTILE'));
        check6 := TDBCheckBox(GetControl('PGH_ASSMUTILE'));
{$ELSE}
        check3 := TCheckBox(GetControl('PGH_ACCTRAV'));
        check4 := TCheckBox(GetControl('PGH_INVPEN'));
        check5 := TCheckBox(GetControl('PGH_MUTILE'));
        check6 := TCheckBox(GetControl('PGH_ASSMUTILE'));
{$ENDIF}
        if check6 <> nil then Check6.onclick := Autreclick;
        if check5 <> nil then Check5.onclick := Autreclick;
        if check3 <> nil then Check3.onclick := Autreclick;
        if check4 <> nil then Check4.onclick := Autreclick;
end ;

procedure TOM_HANDICAPE.affichenom(Sender: TObject);
var
   QSAL : TQuery;
   nomsalarie,salarie : string;
begin
        if sender = nil then exit;
        salarie := GetField('PGH_SALARIE');
        QSAL := OpenSQL('SELECT PSA_LIBELLE FROM SALARIES WHERE PSA_SALARIE="'+salarie+'"',TRUE);
        Nomsalarie:='';
        If not QSAL.eof then nomsalarie := QSAL.FindField('PSA_LIBELLE').AsString;    // // PortageCWAS
        Ferme(QSAL);
        Ecran.caption := 'Saisie travailleurs handicapés de '+nomsalarie;
end;

procedure TOM_HANDICAPE.Autreclick(Sender: TObject);
begin
//Cette procédure concerne l'incompatibilité entre les champs accident du travail,mutié de guerre,.....
        if sender=nil then exit;
        if GetControlText('PGH_HANDICAP') = '-' then exit;
{$IFNDEF EAGLCLIENT}
        if TDBCheckBox(sender).Name = 'PGH_ACCTRAV' then
{$ELSE}
        if TCheckBox(sender).Name = 'PGH_ACCTRAV' then
{$ENDIF}
        begin
                if GetControlText('PGH_ACCTRAV') = 'X' then
                begin
                        SetControlEnabled('PGH_INVPEN',False);
                        SetControlEnabled('PGH_MUTILE',False);
                        SetControlEnabled('PGH_ASSMUTILE',False);
                        SetControlChecked('PGH_INVPEN',False);
                        SetControlChecked('PGH_MUTILE',False);
                        SetControlChecked('PGH_ASSMUTILE',False);
                        SetControlEnabled('PGH_TXINV',True);
                end
                Else
                Begin
                        SetControlEnabled('PGH_INVPEN',True);
                        SetControlEnabled('PGH_MUTILE',True);
                        SetControlEnabled('PGH_ASSMUTILE',True);
                        SetControlEnabled('PGH_TXINV',False);
                        SetControlText('PGH_TXINV','');
                end;
         end;
{$IFNDEF EAGLCLIENT}
        if TDBCheckBox(sender).Name = 'PGH_INVPEN' then
{$ELSE}
        if TCheckBox(sender).Name = 'PGH_INVPEN' then
{$ENDIF}
        begin
                if GetControlText('PGH_INVPEN') = 'X' then
                begin
                        SetControlEnabled('PGH_ACCTRAV',False);
                        SetControlEnabled('PGH_MUTILE',False);
                        SetControlEnabled('PGH_ASSMUTILE',False);
                        SetControlChecked('PGH_ACCTRAV',False);
                        SetControlChecked('PGH_MUTILE',False);
                        SetControlChecked('PGH_ASSMUTILE',False);
                        SetControlEnabled('PGH_TXINV',False);
                        SetControlText('PGH_TXINV','');
                end
                Else
                Begin
                        SetControlEnabled('PGH_ACCTRAV',True);
                        SetControlEnabled('PGH_MUTILE',True);
                        SetControlEnabled('PGH_ASSMUTILE',True);
                end;
        end;
{$IFNDEF EAGLCLIENT}
        if TDBCheckBox(sender).Name = 'PGH_MUTILE' then
{$ELSE}
        if TCheckBox(sender).Name = 'PGH_MUTILE' then
{$ENDIF}
        begin
                if GetControlText('PGH_MUTILE') = 'X' then
                begin
                        SetControlEnabled('PGH_INVPEN',False);
                        SetControlEnabled('PGH_ACCTRAV',False);
                        SetControlEnabled('PGH_ASSMUTILE',False);
                        SetControlChecked('PGH_INVPEN',False);
                        SetControlChecked('PGH_ACCTRAV',False);
                        SetControlChecked('PGH_ASSMUTILE',False);
                        SetControlEnabled('PGH_TXINV',False);
                        SetControlText('PGH_TXINV','');
                end
                Else
                Begin
                        SetControlEnabled('PGH_INVPEN',True);
                        SetControlEnabled('PGH_ACCTRAV',True);
                        SetControlEnabled('PGH_ASSMUTILE',True);
               end;
        end;

{$IFNDEF EAGLCLIENT}
        if TDBCheckBox(sender).Name = 'PGH_ASSMUTILE' then
{$ELSE}
        if TCheckBox(sender).Name = 'PGH_ASSMUTILE' then
{$ENDIF}
        begin
                if GetControlText('PGH_ASSMUTILE') = 'X' then
                begin
                        SetControlEnabled('PGH_INVPEN',False);
                        SetControlEnabled('PGH_MUTILE',False);
                        SetControlEnabled('PGH_ACCTRAV',False);
                        SetControlChecked('PGH_INVPEN',False);
                        SetControlChecked('PGH_MUTILE',False);
                        SetControlChecked('PGH_ACCTRAV',False);
                        SetControlText('PGH_TXINV','');
                end
                Else
                Begin
                        SetControlEnabled('PGH_INVPEN',True);
                        SetControlEnabled('PGH_MUTILE',True);
                        SetControlEnabled('PGH_ACCTRAV',True);
               end;
        end;
end;

procedure TOM_HANDICAPE.Handiclick(Sender: TObject);
begin
        //Cette procédure active les champs handicapés
        if sender=nil then exit;
{$IFNDEF EAGLCLIENT}
        if TDBCheckBox(sender).Name = 'PGH_HANDICAP' then
{$ELSE}
        if TCheckBox(sender).Name = 'PGH_HANDICAP' then
{$ENDIF}
        begin
                if GetControlText('PGH_HANDICAP') = 'X' then
                Begin
                        SetControlEnabled('PGH_ACCTRAV',True);
                        SetControlEnabled('PGH_INVPEN',True);
                        SetControlEnabled('PGH_MUTILE',True);
                        SetControlEnabled('PGH_ASSMUTILE',True);
                        //PT6 Supp des setcontrolenabled a false
                end
                Else
                begin
                        SetControlEnabled('PGH_TXINV',False);
                        SetControlText('PGH_TXINV','');
                        SetControlEnabled('PGH_ACCTRAV',False);
                        SetControlEnabled('PGH_INVPEN',False);
                        SetControlEnabled('PGH_MUTILE',False);
                        SetControlEnabled('PGH_ASSMUTILE',False);
                        SetControlChecked('PGH_ACCTRAV',False);
                        SetControlChecked('PGH_INVPEN',False);
                        SetControlChecked('PGH_MUTILE',False);
                        SetControlChecked('PGH_ASSMUTILE',False);
                end;
        end;
end;

procedure TOM_HANDICAPE.COTOREPClick(Sender: TObject);
begin
        //Cette procédure rend inaccessible la partie handicapé quand on choisi COTOREP et active les champs de la COTOREP
        if sender=nil then exit;
{$IFNDEF EAGLCLIENT}
        if TDBCheckBox(sender).Name = 'PGH_COTOREP' then
{$ELSE}
        if TCheckBox(sender).Name = 'PGH_COTOREP' then
{$ENDIF}
        begin
                if GetControlText('PGH_COTOREP') = 'X' then
                begin
                        SetControlEnabled('PGH_CATEGHANDI',True);
                        SetControlEnabled('PGH_PLACEANT',True);
                        SetControlEnabled('PGH_DATEFINCOT',True);  //PT- 2
                        SetControlEnabled('PGH_FORMATION',True);
                        SetField('PGH_DATEFINCOT',date);     //PT- 2
                        //PT6 Supp des setcontrolenabled a false
                end
                Else
                begin
                        SetControlEnabled('PGH_CATEGHANDI',False);
                        SetControlEnabled('PGH_PLACEANT',False);
                        SetControlEnabled('PGH_DATEFINCOT',False);  //PT- 2
                        // PT4 17/12/2002 PH V591 Toutes les dates sont initialisées à idate1900 ou 2099 au lieu de null
                        SetField('PGH_DATEFINCOT',IDate1900);     //PT- 2
                end;
        end;
end;

 //PT- 1
procedure TOM_HANDICAPE.ExitEdit(Sender: TObject);
var edit : thedit;
    QSAL : TQuery;
    NomSalarie:String;
begin
        edit:=THEdit(Sender);
        if edit <> nil then	//AffectDefautCode que si gestion du code salarié en Numérique
        if (VH_Paie.PgTypeNumSal = 'NUM') and (length(Edit.text) < 11) and (isnumeric(edit.text)) then
        edit.text:=AffectDefautCode(edit,10);
        QSAL := OpenSQL('SELECT PSA_LIBELLE FROM SALARIES WHERE PSA_SALARIE="'+edit.text+'"',TRUE);
        Nomsalarie := '';
        if not QSAL.eof then nomsalarie := QSAL.FindField('PSA_LIBELLE').AsString;   // // PortageCWAS
        Ferme(QSAL);
        Ecran.caption := 'Saisie travailleurs handicapés de '+nomsalarie;
end;


Initialization
  registerclasses ( [ TOM_HANDICAPE ] ) ;
end.
