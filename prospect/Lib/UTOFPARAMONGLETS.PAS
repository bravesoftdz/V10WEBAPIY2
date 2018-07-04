{***********UNITE*************************************************
Auteur  ...... :  MNG
Créé le ...... : 14/05/2001
Modifié le ... :   /  /
Description .. : Source TOF de la TABLE : PARAMONGLETS ()
Mots clefs ... : TOF;PARAMONGLETS
*****************************************************************}
Unit UTOFPARAMONGLETS ;

Interface

Uses StdCtrls, Controls, Classes, sysutils, EntRT, UTOB,
{$IFDEF EAGLCLIENT}
      MainEAGL,
{$ELSE}
      db, {$IFNDEF DBXPRESS}dbtables{BDE},{$ELSE}uDbxDataSet,{$ENDIF} Fe_Main,
{$ENDIF}
     UtilPgi,
     HCtrls, HEnt1, HMsgBox, UTOF,Vierge ;

Const
   CodeFichierProspect : String = '0';
   CodeFichierAction : String = '1';
   CodeFichierOperation : String = '2';
   CodeFichierFournisseur : String = '3';
   CodeFichierArticle : String = '4';
   CodeFichierLigne : String = '8';
   CodeFichierArticleParc : String = 'B';
   CodeFichierParc : String = '5';
   CodeFichierElementParc : String = '9';
   CodeFichierPiece : String = 'D';
   CodeFichierProposition : String = 'V';
   TabletteProspect : String = 'RTLIBCHAMPSLIBRES';
   TabletteAutres : String = 'RTLIBCHAMPS';
   NomTabletteCombo : String = 'RTRPRLIBTABLE';
   ChampsTiers : String = 'T_JURIDIQUE;TTFORMEJURIDIQUE;T_PAYS;TTPAYS;T_ZONECOM;GCZONECOM;T_PARTICULIER;T_NATIONALITE;YYNATIONALITE;T_LANGUE;TTLANGUE;T_SEXE;PGSEXE;T_REPRESENTANT;GCCOMMERCIAL';
   ChampsFou : String = 'T_JURIDIQUE;TTFORMEJURIDIQUE;T_PAYS;TTPAYS;T_NATIONALITE;YYNATIONALITE;T_LANGUE;TTLANGUE';
Type
  TOF_PARAMONGLETS = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnClose                  ; override ;
     Private
         Tablette,TabletteCombo : String;
         NumOnglet: Integer;
         ValChamp,TypeChamp,NomChamp,SavArgument,CodeFichier : String;
         AfficheTout : TCheckBox;
         ListeChamps: THValComboBox;
         ExisteOnglet : Boolean;
         procedure AfficheToutClick(Sender: TObject);
         procedure AfficheOnglet;
         procedure AffichePasOnglet;
         procedure ListeChampsClick(Sender: TObject);
         procedure NomOngletExit(Sender: TObject);
         function ConstruitRetour : String;         
  end ;

Function RTLanceFiche_PARAMONGLETS(Nat,Cod : String ; Range,Lequel,Argument : string) : string;

implementation
uses 
   CbpMCD
   ,CbpEnumerator;


Function RTLanceFiche_PARAMONGLETS(Nat,Cod : String ; Range,Lequel,Argument : string) : string;
begin
  VH_RT.TobChampsPro.Load;

  Result:=AGLLanceFiche(Nat,Cod,Range,Lequel,Argument);
end;

procedure TOF_PARAMONGLETS.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_PARAMONGLETS.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_PARAMONGLETS.OnUpdate ;
begin
  Inherited ;
   TFVierge(ecran).retour:=ConstruitRetour;
end ;

function TOF_PARAMONGLETS.ConstruitRetour : String;
Var Valeur,Nom : String;
begin
  Nom:=THEdit(GetControl('NOMONGLET')).Text;
  if AfficheTout.checked then
     begin
     Valeur:='';
     TypeChamp:='';
     ListeChamps.Value:='';
     end
  else
     begin
     if TypeChamp <> 'BOOLEAN' then
        Valeur:=THMultiValComboBox(GetControl('CRITERECOMBO')).Text
     else
        if TCheckBox(GetControl('CRITEREBOOLEEN')).Checked then Valeur:='X' else Valeur:='-';
     end;
    // 1 : Nom du Champ, 2 : Type du Champ, 3 : Valeur paramétrée 4: Intitulé Onglet
   Result:=ListeChamps.Value+'|'+TypeChamp+'|'+Valeur+'|'+
      Nom+'|';
end;

procedure TOF_PARAMONGLETS.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_PARAMONGLETS.OnArgument (S : String ) ;
var Critere,ValMul : string;
    i,NumChamp,j,iTable,iChamp : integer;
    ListeChamp,ListeLibelle : TStrings;
    TobFille,TobChampsProFille : TOB;
Mcd : IMCDServiceCOM;
Table     : ITableCOM ;
FieldList : IEnumerator ;

begin
  Inherited ;
  MCD := TMCD.GetMcd;
  if not mcd.loaded then mcd.WaitLoaded();
  //
  ValChamp:=''; NomChamp:=''; TypeChamp:='';
  ExisteOnglet:=False;
  NumChamp:= -1;
  i:=0;
  // 0: no onglet, 1 : Nom du Champ, 2 : Type du Champ, 3 : Valeur paramétrée 4: Intitulé Onglet
  Repeat
      Critere:=ReadTokenPipe(S,'|') ;
      if i=0 then
          begin
          NumOnglet := StrToInt(Critere);
          SavArgument:=S;
          end
      else
      if i=1 then NomChamp:=Critere else
      if i=2 then TypeChamp:=Critere else
      if i=3 then ValChamp:=Critere else
      if i=4 then THEdit(GetControl('NOMONGLET')).Text:=Critere else
      if i=5 then CodeFichier:=Critere;
      inc(i);
  until  i=6;
  TabletteCombo:=NomTabletteCombo ;
{$ifdef GIGI}
  if CodeFichier='*' then
    begin  //mcd 18/07/06 passe * au lieu de 0 si catalogue pour en pas voir champ OuiNOn qui ne sert à rien ds ce cas
    CodeFichier:='0';
    SetControlEnabled('OUINON',false);
    end;
{$endif}

  if CodeFichier = CodeFichierProspect then
   Tablette:=TabletteProspect
  else
      begin
      Tablette:=TabletteAutres+'00'+CodeFichier;
      TabletteCombo:=copy(NomTabletteCombo,1,2)+'00'+Codefichier+copy(NomTabletteCombo,6,13) ;
      end;

  if NomChamp<>'' then ExisteOnglet:=True;

   AfficheTout:= TCheckBox(GetControl('OUINON'));
   AfficheTout.OnClick := AfficheToutClick;

   AfficheTout.Caption:=AfficheTout.Caption+' '+THEdit(GetControl('NOMONGLET')).Text;

   ListeChamps:= THValComboBox(GetControl('LISTECHAMPS'));
   ListeChamps.OnClick := ListeChampsClick;

   ListeChamp := TStringList.Create ;
   ListeLibelle := TStringList.Create;

   THEdit(GetControl('NOMONGLET')).Onexit := NomOngletExit;

    //TobChampsProFille:=VH_RT.TobChampsPro.detail[StrToInt(CodeFichier)];
    TobChampsProFille:=VH_RT.TobChampsPro.FindFirst(['CO_CODE'], [CodeFichier], TRUE);
    for j := 0 to TobChampsProFille.Detail.Count-1 do
       begin
          TobFille := TobChampsProFille.Detail[j];
          if ( TobFille.getValue('RDE_ONGLET') <> NumOnglet ) and
             ( TobFille.getValue('RDE_PANEL') <> 99 ) and
             ( (TobFille.getValue('RDE_TYPECHAMP') = 'COMBO') or
               (TobFille.getValue('RDE_TYPECHAMP') = 'BOOLEAN') ) then
             begin
             if TobFille.getValue('RDE_TYPECHAMP') = 'COMBO' then
                begin
                i:=StrToInt(copy(TobFille.getValue('RDE_NOMCHAMP'),16,length(TobFille.getValue('RDE_NOMCHAMP'))));
                if ( i < 11 ) then
                   ListeLibelle.add(RechDom(Tablette,'CL'+IntToStr(i),FALSE))
                else
                   ListeLibelle.add(RechDom(Tablette,'CL'+Chr(i+55),FALSE));
                end
             else
                begin
                i:=StrToInt(copy(TobFille.getValue('RDE_NOMCHAMP'),15,length(TobFille.getValue('RDE_NOMCHAMP'))));
                ListeLibelle.add(RechDom(Tablette,'BL'+IntToStr(i),FALSE));
                end;
             ListeChamp.add(TobFille.getValue('RDE_NOMCHAMP'));
             if NomChamp = TobFille.getValue('RDE_NOMCHAMP') then
                NumChamp:= ListeChamp.Count;
             end;
       end;
    if CodeFichier = CodeFichierProspect then
        begin
        for i:=1 to 10 do
           begin
           if i < 10 then
              ValMul:=RechDomZoneLibre ('CT'+IntToStr(i),FALSE)
           else
              ValMul:=RechDomZoneLibre ('CTA',FALSE);
           if copy(ValMul,1,1) <> '.' then
               begin
               ListeLibelle.add(ValMul);
               if i < 10 then
                  Critere:= 'YTC_TABLELIBRETIERS'+IntToStr(i)
               else
                  Critere:= 'YTC_TABLELIBRETIERSA';
               ListeChamp.add(Critere);
               if NomChamp = Critere then
                    NumChamp:= ListeChamp.Count;
               end;
           end;
        // champs suplémentaires de la fiche tiers
          table := Mcd.getTable(mcd.PrefixetoTable('T'));
          FieldList := Table.Fields;
          FieldList.Reset();
          While FieldList.MoveNext do
            begin
               if Pos((FieldList.Current as IFieldCOM).name,ChampsTiers) <> 0 then
                   begin
                   ListeLibelle.add((FieldList.Current as IFieldCOM).libelle);
                   ListeChamp.add((FieldList.Current as IFieldCOM).name);
                   if NomChamp = (FieldList.Current as IFieldCOM).name then
                      NumChamp:= ListeChamp.Count;
                   end;
            end;
        end
    else
    if CodeFichier = CodeFichierFournisseur then
        begin
        for i:=1 to 3 do
           begin
           ValMul:=RechDomZoneLibre ('FT'+IntToStr(i),FALSE);
           if copy(ValMul,1,1) <> '.' then
               begin
               ListeLibelle.add(ValMul);
               Critere:= 'YTC_TABLELIBREFOU'+IntToStr(i);
               ListeChamp.add(Critere);
               if NomChamp = Critere then
                    NumChamp:= ListeChamp.Count;
               end;
           end;
        // champs suplémentaires de la fiche tiers
          table := Mcd.getTable(mcd.PrefixetoTable('T'));
          FieldList := Table.Fields;
          FieldList.Reset();
          While FieldList.MoveNext do
            begin
               if Pos((FieldList.Current as IFieldCOM).name,ChampsFou) <> 0 then
                   begin
                   ListeLibelle.add((FieldList.Current as IFieldCOM).libelle);
                   ListeChamp.add((FieldList.Current as IFieldCOM).name);
                   if NomChamp = (FieldList.Current as IFieldCOM).name then
                      NumChamp:= ListeChamp.Count;
                   end;
            end;
        end
    else
    if CodeFichier = CodeFichierArticle then
       begin
       ListeLibelle.add(TraduireMemoire('Type d''article'));
       Critere:= 'GA_TYPEARTICLE';
       ListeChamp.add(Critere);
       if NomChamp = Critere then
            NumChamp:= ListeChamp.Count;
       for i:=1 to 10 do
           begin
           if i < 10 then
              ValMul:=RechDomZoneLibre ('AT'+IntToStr(i),FALSE)
           else
              ValMul:=RechDomZoneLibre ('ATA',FALSE);
           if copy(ValMul,1,1) <> '.' then
               begin
               ListeLibelle.add(ValMul);
               if i < 10 then
                  Critere:= 'GA_LIBREART'+IntToStr(i)
               else
                  Critere:= 'GA_LIBREARTA';
               ListeChamp.add(Critere);
               if NomChamp = Critere then
                    NumChamp:= ListeChamp.Count;
               end;
           end;
         for i:=1 to 3 do
             begin
             ValMul:=RechDom('GCLIBFAMILLE','LF'+IntToStr(i),FALSE) ;
             if copy(ValMul,1,2) <> '.-' then
                 begin
                 ListeLibelle.add(ValMul);
                 Critere:= 'GA_FAMILLENIV'+IntToStr(i) ;
                 ListeChamp.add(Critere);
                 if NomChamp = Critere then
                      NumChamp:= ListeChamp.Count;
                 end;
             end;
       //TFVierge(Ecran).HelpContext:=111000256;
       end
    else
      if CodeFichier = CodeFichierLigne then
         begin
         ListeLibelle.add(TraduireMemoire('Nature de pièce'));
         Critere:= 'GL_NATUREPIECEG';
         ListeChamp.add(Critere);
         if NomChamp = Critere then
              NumChamp:= ListeChamp.Count;
         ListeLibelle.add(TraduireMemoire('Type d''article'));
         Critere:= 'GL_TYPEARTICLE';
         ListeChamp.add(Critere);
         if NomChamp = Critere then
              NumChamp:= ListeChamp.Count;

         for i:=1 to 10 do
             begin
             if i < 10 then
                ValMul:=RechDomZoneLibre ('AT'+IntToStr(i),FALSE)
             else
                ValMul:=RechDomZoneLibre ('ATA',FALSE);
             if copy(ValMul,1,1) <> '.' then
                 begin
                 ListeLibelle.add(ValMul);
                 if i < 10 then
                    Critere:= 'GL_LIBREART'+IntToStr(i)
                 else
                    Critere:= 'GL_LIBREARTA';
                 ListeChamp.add(Critere);
                 if NomChamp = Critere then
                      NumChamp:= ListeChamp.Count;
                 end;
             end;
           for i:=1 to 3 do
               begin
               ValMul:=RechDom('GCLIBFAMILLE','LF'+IntToStr(i),FALSE) ;
               if copy(ValMul,1,2) <> '.-' then
                   begin
                   ListeLibelle.add(ValMul);
                   Critere:= 'GL_FAMILLENIV'+IntToStr(i) ;
                   ListeChamp.add(Critere);
                   if NomChamp = Critere then
                        NumChamp:= ListeChamp.Count;
                   end;
               end;
         //TFVierge(Ecran).HelpContext:=111000256;
         end
      else
      if CodeFichier = CodeFichierPiece then
         begin
         ListeLibelle.add('Nature de pièce');
         Critere:= 'GP_NATUREPIECEG';
         ListeChamp.add(Critere);
         if NomChamp = Critere then
              NumChamp:= ListeChamp.Count;
         for i:=1 to 10 do
             begin
             if i < 10 then
                ValMul:=RechDomZoneLibre ('CT'+IntToStr(i),FALSE)
             else
                ValMul:=RechDomZoneLibre ('CTA',FALSE);
             if copy(ValMul,1,1) <> '.' then
                 begin
                 ListeLibelle.add(ValMul);
                 if i < 10 then
                    Critere:= 'GP_LIBRETIERS'+IntToStr(i)
                 else
                    Critere:= 'GP_LIBRETIERS';
                 ListeChamp.add(Critere);
                 if NomChamp = Critere then
                      NumChamp:= ListeChamp.Count;
                 end;
             end;
         end
      else
        if CodeFichier = CodeFichierAction then
           begin
           for i:=1 to 3 do
               begin
               ValMul:=RechDom('RTLIBCLACTPERSP','AL'+IntToStr(i),FALSE) ;
               if copy(ValMul,1,2) <> '.-' then
                   begin
                   ListeLibelle.add(ValMul);
                   Critere:= 'RAC_TABLELIBRE'+IntToStr(i) ;
                   ListeChamp.add(Critere);
                   if NomChamp = Critere then
                        NumChamp:= ListeChamp.Count;
                   end;
               end;
           ListeLibelle.add(TraduireMemoire('Type d''action'));
           Critere:= 'RAC_TYPEACTION';
           ListeChamp.add(Critere);
           if NomChamp = Critere then
                NumChamp:= ListeChamp.Count;
           TFVierge(Ecran).HelpContext:=111000256;
           end
        else
        if CodeFichier = CodeFichierProposition then
           begin
           for i:=1 to 3 do
               begin
               ValMul:=RechDom('RTLIBCLACTPERSP','PL'+IntToStr(i),FALSE) ;
               if copy(ValMul,1,2) <> '.-' then
                   begin
                   ListeLibelle.add(ValMul);
                   Critere:= 'RPE_TABLELIBREPER'+IntToStr(i) ;
                   ListeChamp.add(Critere);
                   if NomChamp = Critere then
                        NumChamp:= ListeChamp.Count;
                   end;
               end;
           for i:=1 to 5 do
               begin
               ValMul:=RechDom('RTLIBPERSPECTIVE','TL'+IntToStr(i),FALSE) ;
               if copy(ValMul,1,2) <> '.-' then
                   begin
                   ListeLibelle.add(ValMul);
                   Critere:= 'RPE_RPETABLELIBRE'+IntToStr(i) ;
                   ListeChamp.add(Critere);
                   if NomChamp = Critere then
                        NumChamp:= ListeChamp.Count;
                   end;
               end;
           ListeLibelle.add(TraduireMemoire('Situation'));
           Critere:= 'RPE_ETATPER';
           ListeChamp.add(Critere);
           if NomChamp = Critere then
                NumChamp:= ListeChamp.Count;
           ListeLibelle.add(TraduireMemoire('Nature'));
           Critere:= 'RPE_TYPEPERSPECTIV';
           ListeChamp.add(Critere);
           if NomChamp = Critere then
                NumChamp:= ListeChamp.Count;

           TFVierge(Ecran).HelpContext:=111000256;
           end
        else
        if CodeFichier = CodeFichierArticleParc then
           begin
           ListeLibelle.add(TraduireMemoire('Qté indivisible'));
           Critere:= 'WAP_QTEINDIVISIBLE';
           ListeChamp.add(Critere);
           if NomChamp = Critere then
                NumChamp:= ListeChamp.Count;
           for i:=1 to 10 do
               begin
               if i < 10 then
                  ValMul:=RechDomZoneLibre('WT'+IntToStr(i),FALSE)
               else
                  ValMul:=RechDomZoneLibre('WTA',FALSE);
               if copy(ValMul,1,2) <> '.-' then
                   begin
                   ListeLibelle.add(ValMul);
                   if i < 10 then
                      Critere:= 'WAP_LIBREWAP'+IntToStr(i)
                   else
                      Critere:= 'WAP_LIBREWAPA';
                   ListeChamp.add(Critere);
                   if NomChamp = Critere then
                        NumChamp:= ListeChamp.Count;
                   end;
               end;
             for i:=1 to 3 do
                 begin
                 ValMul:=RechDomZoneLibre('WB'+IntToStr(i),FALSE) ;
                 if copy(ValMul,1,2) <> '.-' then
                     begin
                     ListeLibelle.add(ValMul);
                     Critere:= 'WAP_BOOLLIBRE'+IntToStr(i) ;
                     ListeChamp.add(Critere);
                     if NomChamp = Critere then
                          NumChamp:= ListeChamp.Count;
                     end;
                 end;
           //TFVierge(Ecran).HelpContext:=111000256;
           end
        else
          if CodeFichier = CodeFichierParc then
             begin
             ListeLibelle.add(TraduireMemoire('Etat parc'));
             Critere:= 'WPC_ETATPARC';
             ListeChamp.add(Critere);
             if NomChamp = Critere then
                  NumChamp:= ListeChamp.Count;
             for i:=1 to 10 do
                 begin
                 if i < 10 then
                    ValMul:=RechDom('WLIBELLELIBREWPC','TWPC_LIBREWPC'+IntToStr(i),FALSE)
                 else
                    ValMul:=RechDom('WLIBELLELIBREWPC','TWPC_LIBREWPCA',FALSE);
                 if copy(ValMul,1,1) <> '.-' then
                     begin
                     ListeLibelle.add(ValMul);
                     if i < 10 then
                        Critere:= 'WPC_LIBREWPC'+IntToStr(i)
                     else
                        Critere:= 'WPC_LIBREWPCA';
                     ListeChamp.add(Critere);
                     if NomChamp = Critere then
                          NumChamp:= ListeChamp.Count;
                     end;
                 end;
               for i:=1 to 3 do
                   begin
                   ValMul:=RechDom('WBOOLLIBREWPC','TWPC_BOOLLIBRE'+IntToStr(i),FALSE) ;
                   if copy(ValMul,1,2) <> '.-' then
                       begin
                       ListeLibelle.add(ValMul);
                       Critere:= 'WPC_BOOLLIBRE'+IntToStr(i) ;
                       ListeChamp.add(Critere);
                       if NomChamp = Critere then
                            NumChamp:= ListeChamp.Count;
                       end;
                   end;
             //TFVierge(Ecran).HelpContext:=111000256;
             end
          else
            if CodeFichier = CodeFichierElementParc then
               begin
               ListeLibelle.add(TraduireMemoire('Etat parc'));
               Critere:= 'WPN_ETATPARC';
               ListeChamp.add(Critere);
               if NomChamp = Critere then
                    NumChamp:= ListeChamp.Count;
               for i:=1 to 10 do
                   begin
                   if i < 10 then
                      ValMul:=RechDom('WLIBELLELIBREWPN','TWPN_LIBREWPN'+IntToStr(i),FALSE)
                   else
                      ValMul:=RechDom('WLIBELLELIBREWPN','TWPN_LIBREWPNA',FALSE);
                   if copy(ValMul,1,2) <> '.-' then
                       begin
                       ListeLibelle.add(ValMul);
                       if i < 10 then
                          Critere:= 'WPN_LIBREWPN'+IntToStr(i)
                       else
                          Critere:= 'WPN_LIBREWPNA';
                       ListeChamp.add(Critere);
                       if NomChamp = Critere then
                            NumChamp:= ListeChamp.Count;
                       end;
                   end;
                 for i:=1 to 3 do
                     begin
                     ValMul:=RechDom('WBOOLLIBREWPN','TWPN_BOOLLIBRE'+IntToStr(i),FALSE) ;
                     if copy(ValMul,1,2) <> '.-' then
                         begin
                         ListeLibelle.add(ValMul);
                         Critere:= 'WPN_BOOLLIBRE'+IntToStr(i) ;
                         ListeChamp.add(Critere);
                         if NomChamp = Critere then
                              NumChamp:= ListeChamp.Count;
                         end;
                     end;
               //TFVierge(Ecran).HelpContext:=111000256;
               end
            else
            if CodeFichier = CodeFichierOperation then
               TFVierge(Ecran).HelpContext:=111000215
            else
            if CodeFichier = CodeFichierFournisseur then
               TFVierge(Ecran).HelpContext:=111000201
            else
            if CodeFichier = CodeFichierProposition then
               TFVierge(Ecran).HelpContext:=111000263
            else
               TFVierge(Ecran).HelpContext:=0;

    THValcomboBox(GetControl('LISTECHAMPS')).Items.Assign(ListeLibelle);
    THValcomboBox(GetControl('LISTECHAMPS')).Values.Assign(ListeChamp);
    if (ExisteOnglet) and (NumChamp <> -1) then
        THValcomboBox(GetControl('LISTECHAMPS')).ItemIndex := NumChamp-1;
    ListeChamp.free;ListeLibelle.Free;

   if (not ExisteOnglet) or (NomChamp='') then
      begin
      AfficheTout.checked:=True;
      AfficheOnglet;
      end
   else
      begin
      AfficheTout.checked:=False;
      AffichePasOnglet;
      end;
end ;

procedure TOF_PARAMONGLETS.OnClose ;
var Retour,Ret : string;
begin
  Inherited ;
if TFVierge(ecran).retour = '' then
    begin
    Retour:=ConstruitRetour;
    Ret:=Retour+CodeFichier;
    if Ret <> SavArgument then
        begin
        Case PGIAskCancel('Voulez-vous enregistrer les modifications ?',TFVierge(ecran).Caption) of
          mrNo : Retour:=SavArgument;
          mrCancel : LastError:=1;
        end ;
        end;
    TFVierge(ecran).retour:=Retour;
    end;
end ;

procedure TOF_PARAMONGLETS.AfficheToutClick(Sender: TObject);
begin
if AfficheTout.Checked then
   AfficheOnglet
else
   begin
   SetControlVisible('CONDITION',True);
   SetControlVisible('LISTE',True);
   SetControlVisible('LISTECHAMPS',True);
   end;
end;

procedure TOF_PARAMONGLETS.NomOngletExit(Sender: TObject);
var Nom : string;
begin
Nom:=copy(AfficheTout.Caption,1,34)+' '+THEdit(GetControl('NOMONGLET')).Text;
AfficheTout.Caption:=Nom;
end;

procedure TOF_PARAMONGLETS.ListeChampsClick(Sender: TObject);
var Q : TQuery;
begin
// rechercher le type du champ sélectionné
if ListeChamps.Value <> '' then
    begin
    Q := OpenSQL ('SELECT DH_TYPECHAMP from DECHAMPS where DH_NOMCHAMP="'+ListeChamps.Value+'"',True) ;
    if not Q.EOF then
       begin
       TypeChamp:=Q.Fields[0].AsString;
       if copy(TypeChamp,1,7) = 'VARCHAR' then TypeChamp:='COMBO';
       ValChamp:='';
       AffichePasOnglet;
       end;
    Ferme(Q) ;
    end
else
    AfficheOnglet;
end;

procedure TOF_PARAMONGLETS.AfficheOnglet;
begin
SetControlVisible('CRITEREBOOLEEN',FALSE) ;

SetControlVisible('NOMCOMBO',FALSE);
SetControlVisible('CONDITION',False);
SetControlVisible('LISTE',FALSE);
SetControlText('LISTECHAMPS','');
SetControlVisible('LISTECHAMPS',FALSE);
SetControlText('CRITERECOMBO','');
SetControlVisible('CRITERECOMBO',FALSE);
ValChamp:='';
end;

procedure TOF_PARAMONGLETS.AffichePasOnglet;
Var i : integer;
    Tablettes,NomTablette : String;
begin
SetControlVisible('CONDITION',True);
SetControlVisible('LISTE',True);
SetControlVisible('LISTECHAMPS',True);

if TypeChamp<>'BOOLEAN' then
   begin
   SetControlVisible('CRITEREBOOLEEN',FALSE) ;
   SetControlVisible('CRITERECOMBO',True) ;
   SetControlVisible('NOMCOMBO',True) ;
   SetControlText('NOMCOMBO',ListeChamps.Text);
   SetControlText('CRITERECOMBO',ValChamp);
   if copy(ListeChamps.Value,1,3) = 'RAC' then
      begin
      if copy(ListeChamps.Value,5,3) <> 'TAB' then
        SetControlProperty('CRITERECOMBO','DataType','RTTYPEACTION')
      else
          begin
          i:=StrToInt(copy(ListeChamps.Value,15,15));
          SetControlProperty('CRITERECOMBO','DataType','RTRPRLIBACTION'+IntToStr(i));
          end;
      end
   else
   if copy(ListeChamps.Value,1,3) = 'RPE' then
      begin
      if ListeChamps.Value = 'RPE_ETATPER' then
        SetControlProperty('CRITERECOMBO','DataType','RTETATPERSPECTIVE')
      else
      if ListeChamps.Value = 'RPE_TYPEPERSPECTIV' then
        SetControlProperty('CRITERECOMBO','DataType','RTTYPEPERSPECTIVE')
      else
      if copy(ListeChamps.Value,5,3) = 'RPE' then
          begin
          i:=StrToInt(copy(ListeChamps.Value,18,18));
          SetControlProperty('CRITERECOMBO','DataType','RTRRPETABLELIBRE'+IntToStr(i));
          end
      else
        if copy(ListeChamps.Value,5,3) = 'TAB' then
          begin
          i:=StrToInt(copy(ListeChamps.Value,18,18));
          SetControlProperty('CRITERECOMBO','DataType','RTRPRLIBPERSPECTIVE'+IntToStr(i));
          end;
      end
   else
     if copy(ListeChamps.Value,1,1) = 'T' then
        begin
        i:=Pos(ListeChamps.Value,ChampsTiers);
        Tablettes:=Copy(ChampsTiers,i,Length(ChampsTiers));
        NomTablette:=ReadToKenSt(Tablettes);
        NomTablette:=ReadToKenSt(Tablettes);
        SetControlProperty('CRITERECOMBO','DataType',NomTablette);
        end
     else
       if (copy(ListeChamps.Value,1,2) = 'GA') or (copy(ListeChamps.Value,1,2) = 'GL') then
          begin
          if copy(ListeChamps.Value,1,15) = 'GL_NATUREPIECEG' then
             begin
             SetControlProperty('CRITERECOMBO','DataType','GCNATUREPIECEG');
             SetControlProperty('CRITERECOMBO','Plus','AND (GPP_INFOSCOMPL="X")');
             end
          else
            if copy(ListeChamps.Value,4,10) = 'FAMILLENIV' then
              begin
              i:=StrToInt(copy(ListeChamps.Value,14,1));
              SetControlProperty('CRITERECOMBO','DataType','GCFAMILLENIV'+IntToStr(i));
              end
            else
              if copy(ListeChamps.Value,4,8) = 'LIBREART' then
                begin
                i:=StrToInt(copy(ListeChamps.Value,12,1));
                SetControlProperty('CRITERECOMBO','DataType','GCLIBREART'+IntToStr(i));
                end
              else if copy(ListeChamps.Value,4,11) = 'TYPEARTICLE' then
                      SetControlProperty('CRITERECOMBO','DataType','GCTYPEARTICLE');
          end
       else
         if (copy(ListeChamps.Value,1,3) = 'WAP') then
            begin
            if copy(ListeChamps.Value,5,8) = 'LIBREWAP' then
              begin
              i:=StrToInt(copy(ListeChamps.Value,13,1));
              SetControlProperty('CRITERECOMBO','DataType','WLIBREWAP'+IntToStr(i));
              end;
            end
          else
           if (copy(ListeChamps.Value,1,3) = 'WPC') then
              begin
              if copy(ListeChamps.Value,5,8) = 'LIBREWPC' then
                begin
                i:=StrToInt(copy(ListeChamps.Value,13,1));
                SetControlProperty('CRITERECOMBO','DataType','WLIBREWPC'+IntToStr(i));
                end
              else if copy(ListeChamps.Value,5,8) = 'ETATPARC' then
                      SetControlProperty('CRITERECOMBO','DataType','WETATPARC');
              end
            else
             if (copy(ListeChamps.Value,1,3) = 'WPN') then
                begin
                if copy(ListeChamps.Value,5,8) = 'LIBREWPN' then
                  begin
                  i:=StrToInt(copy(ListeChamps.Value,13,1));
                  SetControlProperty('CRITERECOMBO','DataType','WLIBREWPN'+IntToStr(i));
                  end
                else if copy(ListeChamps.Value,5,8) = 'ETATPARC' then
                        SetControlProperty('CRITERECOMBO','DataType','WETATPARC');
                end
             else
               if copy(ListeChamps.Value,1,1) = 'Y' then
                  begin
                  if CodeFichier=CodeFichierProspect then
                    begin
                    i:=StrToInt(copy(ListeChamps.Value,20,length(ListeChamps.Value)));
                    if i < 10 then
                      SetControlProperty('CRITERECOMBO','DataType','GCLIBRETIERS'+IntToStr(i))
                    else
                      SetControlProperty('CRITERECOMBO','DataType','GCLIBRETIERSA');
                    end
                  else
                    begin
                    i:=StrToInt(copy(ListeChamps.Value,18,length(ListeChamps.Value)));
                    SetControlProperty('CRITERECOMBO','DataType','GCLIBREFOU'+IntToStr(i));
                    end;
                  end
               else
                 if (copy(ListeChamps.Value,1,2) = 'GP')then
                    begin
                    if copy(ListeChamps.Value,1,15) = 'GP_NATUREPIECEG' then
                       begin
                       SetControlProperty('CRITERECOMBO','DataType','GCNATUREPIECEG');
                       SetControlProperty('CRITERECOMBO','Plus','AND (GPP_INFOSCPLPIECE="X")');
                       end
                    else
                        if copy(ListeChamps.Value,4,10) = 'LIBRETIERS' then
                          begin
                          i:=StrToInt(copy(ListeChamps.Value,14,1));
                          SetControlProperty('CRITERECOMBO','DataType','GCLIBRETIERS'+IntToStr(i));
                          end;
                    end
                  else
                    begin
                    i:=StrToInt(copy(ListeChamps.Value,16,length(ListeChamps.Value)));
                    SetControlProperty('CRITERECOMBO','DataType',TabletteCombo+IntToStr(i));
                    end

   end
else
   begin
   SetControlVisible('CRITEREBOOLEEN',True) ;
   SetControlVisible('CRITERECOMBO',FALSE) ;
   SetControlVisible('NOMCOMBO',FALSE) ;
   SetControlCaption('CRITEREBOOLEEN',ListeChamps.Text);
   if ValChamp='X' then
      SetControlChecked ('CRITEREBOOLEEN',True)
   else
      SetControlChecked ('CRITEREBOOLEEN',False);
   end;
end;

Initialization
  registerclasses ( [ TOF_PARAMONGLETS ] ) ;
end.
