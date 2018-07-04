unit utofyymuluserconf_sel ;

interface

Uses StdCtrls,
     Controls,
     Classes,
     menus,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS}dbtables{BDE},{$ELSE}uDbxDataSet,{$ENDIF}
     mul,
     fe_main,
{$ELSE}
     eMul,
     uTob,
     maineagl,
{$ENDIF}
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     HDB,
     UTOF,
     YNewMessage,
     Mailol,
     windows,
     utofzoneslibres
     ;

Type
  TOF_YYMULUSERCONF_SEL = class (TOF)

  public
         procedure OnArgument (strArgs:string); override;
         procedure OnClose;                     override;

  private
         procedure OnOuvrirClick (Sender:TObject);
  end;

function YYLanceMulUserAgenda (strArgs:string):string;


implementation

uses
    HTB97;

function YYLanceMulUserAgenda (strArgs:string):string;
begin
     Result := AGLLanceFiche ('YY','YYMULUSERCONF_SEL', '', '', strArgs);
end;

procedure TOF_YYMULUSERCONF_SEL.OnArgument (strArgs:string);
begin
     TToolBarButton97 (GetControl ('BOUVRIR')).OnClick := OnOuvrirClick;
end;

procedure TOF_YYMULUSERCONF_SEL.OnClose;
begin
end;

procedure TOF_YYMULUSERCONF_SEL.OnOuvrirClick (Sender:TObject);
var
   i          :integer;
   FListe     :THDBGrid;
   strUser    :string;
   strRetour  :string;
begin
     strRetour := '';

     // On renvoie les codes utilisateurs sélectionnés
     FListe := GetControl ('FLISTE') as THDBGrid;
     if FListe <> nil then
     begin
          if FListe.AllSelected = TRUE then
          begin
               // #### manque le FetchLesTous en eaglclient, car sinon "Tout sélectionner"
               // #### se limite à ce qui est déjà chargé dans la grille...
               with Ecran as TFMul do
               begin
                    Q.First;
                    while not Q.EOF do
                    begin
                         strUser := Q.FindField ('US_UTILISATEUR').AsString;
                         if Pos (strUser, strRetour) < 1 then
                         begin
                              if strRetour <> '' then
                                 strRetour := strRetour + ';';
                              strRetour := strRetour + strUser;
                         end;
                         Q.Next;
                    end;
               end;
          end
          else
          begin
               for i := 0 to FListe.NbSelected-1 do
               begin
                    FListe.GotoLeBookmark (i);
            {$IFDEF EAGLCLIENT}
                    TFMul(Ecran).Q.TQ.Seek(FListe.Row - 1);
            {$ELSE}
            {$ENDIF}
                    strUser := GetField ('US_UTILISATEUR');
                    if Pos (strUser, strRetour) < 1 then
                    begin
                         if i > 0 then
                            strRetour := strRetour + ';';
                         strRetour := strRetour + strUser;
                    end;
               end;
          end;
     end;

     // Affectation à la valeur retour du lancement de fiche
     TFMul (Ecran).Retour := strRetour;
     Ecran.Close;
end;


initialization
  RegisterClasses ( [TOF_YYMULUSERCONF_SEL] ) ;

end.

