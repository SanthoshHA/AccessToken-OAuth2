codeunit 50119 "Virgin Integration Mgmt"
{
    TableNo = "Job Queue Entry";
    Permissions = tabledata "Sales Cr.Memo Header" = rimd, tabledata "Sales Invoice Header" = rmid;

    trigger OnRun()
    var
        RecRef: RecordRef;
    begin
        AzadeaSetup.Get();
        AzadeaSetup.TestField("IH Cashier");
        AzadeaSetup.TestField("IH Loyalty");
        AzadeaSetup.TestField("IH Register");
        AzadeaSetup.TestField("IH Store");
        AzadeaSetup.TestField("IH Type");
        AzadeaSetup.TestField("IM Type");

        Rec.TestField("Record ID to Process");
        RecRef.Get(Rec."Record ID to Process");
        if RecRef.Number = Database::"Sales Invoice Header" then
            SendSalesInvoiceToVirgin(RecRef)
        else
            SendSalesReturnToVirgin(RecRef);
    end;

    /// <summary>
    /// Build JSON request for invoice and send to Virgin
    /// </summary>
    /// <param name="RecRef">RecordRef.</param>
    local procedure SendSalesInvoiceToVirgin(RecRef: RecordRef)
    var
        SalesInvHdr: Record "Sales Invoice Header";
        JsonText: Text;
        FreeTextLbl: Label 'Order No. - %1';
    begin
        AzadeaSetup.TestField("Virgin Sales Invoice URL");
        RecRef.SetTable(SalesInvHdr);
        GenerateTRSNumberIfBlank(SalesInvHdr."Integration Refrence No.");
        JsonText := BuildJsonText(SalesInvHdr."No.", SalesInvHdr."Order No.", SalesInvHdr."Posting Date", SalesInvHdr.SystemCreatedAt, true);
        SetJobQueueStatusForInvoice(SalesInvHdr, SalesInvHdr."Virgin API Job Queue Status"::Posting);
        ClearLastError();
        if not PostSalesDataToVirgin(JsonText, AzadeaSetup."Virgin Sales Return URL", StrSubstNo(FreeTextLbl, SalesInvHdr."Order No.")) then begin
            SetJobQueueStatusForInvoice(SalesInvHdr, SalesInvHdr."Virgin API Job Queue Status"::Error);
            Error(GetLastErrorText());
        end;
        SetJobQueueStatusForInvoice(SalesInvHdr, SalesInvHdr."Virgin API Job Queue Status"::" ");
    end;

    /// <summary>
    /// Build JSON request for creditMemo and send to Virgin
    /// </summary>
    /// <param name="RecRef">RecordRef.</param>
    local procedure SendSalesReturnToVirgin(RecRef: RecordRef)
    var
        SalesCreditMemoHdr: Record "Sales Cr.Memo Header";
        JsonText: Text;
        FreeTextLbl: Label 'Return Order No. - %1';
    begin
        AzadeaSetup.TestField("Virgin Sales Return URL");
        RecRef.SetTable(SalesCreditMemoHdr);
        GenerateTRSNumberIfBlank(SalesCreditMemoHdr."Integration Refrence No.");
        JsonText := BuildJsonText(SalesCreditMemoHdr."No.", SalesCreditMemoHdr."Original Order No.", SalesCreditMemoHdr."Posting Date", SalesCreditMemoHdr.SystemCreatedAt, false);
        SetJobQueueStatusForReturn(SalesCreditMemoHdr, SalesCreditMemoHdr."Virgin API Job Queue Status"::Posting);
        ClearLastError();
        if not PostSalesDataToVirgin(JsonText, AzadeaSetup."Virgin Sales Return URL", StrSubstNo(FreeTextLbl, SalesCreditMemoHdr."Return Order No.")) then begin
            SetJobQueueStatusForReturn(SalesCreditMemoHdr, SalesCreditMemoHdr."Virgin API Job Queue Status"::Error);
            Error(GetLastErrorText());
        end;
        SetJobQueueStatusForReturn(SalesCreditMemoHdr, SalesCreditMemoHdr."Virgin API Job Queue Status"::" ");
    end;

    local procedure BuildJsonText(DocNo: Code[20]; OrderNo: Code[20]; PostingDate: Date; CreatedDateTime: DateTime; IsInvoice: Boolean) JsonText: Text
    var
        JSalesHdr: JsonObject;
        JArgs0: JsonObject;
        JItemArray: JsonArray;
        TotalAmt: Decimal;
    begin
        if IsInvoice then
            JItemArray := SalesInvoiceLinesToJson(DocNo, TotalAmt)
        else
            JItemArray := SalesCreditMemoLinesToJson(DocNo, TotalAmt);

        JSalesHdr.Add('IH_CASHIER', AzadeaSetup."IH Cashier");
        JSalesHdr.Add('IH_DATE', Format(PostingDate, 0, '<Year,2><Month,2><Day,2>'));
        JSalesHdr.Add('IH_LOYALTY', AzadeaSetup."IH Loyalty");
        JSalesHdr.Add('IH_ORDER_NUM', OrderNo);
        JSalesHdr.Add('IH_REGISTER', AzadeaSetup."IH Register");
        JSalesHdr.Add('IH_STORE', AzadeaSetup."IH Store");
        JSalesHdr.Add('IH_TIME', ConvertTimeToSeconds(CreatedDateTime));
        JSalesHdr.Add('IH_TOTAL_AMT', TotalAmt);
        JSalesHdr.Add('IH_TRS_NUMBER', TRSNo);
        JSalesHdr.Add('IH_TYPE', AzadeaSetup."IH Type");

        //Items into JSON Array
        JSalesHdr.Add('i_ITEM', JItemArray);

        //MEANSOP into JSON Array
        JSalesHdr.Add('i_MEANSOP', MEANSOPToJson(TotalAmt));

        JArgs0.Add('args0', JSalesHdr);
        JArgs0.WriteTo(JsonText);
    end;

    local procedure BuildItemJson(UnitPrice: Decimal; Quantity: Decimal; BarcodeNo: Code[50]; var JSalesInvLines: JsonArray)
    var
        JSalesInvLine: JsonObject;
    begin
        JSalesInvLine.Add('II_BARCODE', BarcodeNo);
        JSalesInvLine.Add('II_EXT_PRICE', UnitPrice);
        JSalesInvLine.Add('II_OVR_PRICE', UnitPrice);
        JSalesInvLine.Add('II_QUANTITY', Quantity);
        JSalesInvLine.Add('II_UNI_PRICE', UnitPrice);
        JSalesInvLine.Add('II_SERIAL_NB', 'None');
        JSalesInvLines.Add(JSalesInvLine);
    end;

    local procedure MEANSOPToJson(TotalAmt: Decimal): JsonArray
    var
        JMEANSOPLines: JsonArray;
        JMEANSOPLine: JsonObject;
    begin
        JMEANSOPLine.Add('IM_AMOUNT', TotalAmt);
        JMEANSOPLine.Add('IM_TYPE', AzadeaSetup."IM Type");
        JMEANSOPLines.Add(JMEANSOPLine);
        exit(JMEANSOPLines);
    end;

    /// <summary>
    /// Used to post sales invoice or credti memo details to virgin
    /// </summary>
    /// <param name="JsonText">Text.</param>
    /// <param name="URL">Text.</param>
    [TryFunction]
    local procedure PostSalesDataToVirgin(JsonText: Text; URL: Text; FreeText: Text)
    var
        MessageHeaderL: Record "Web Service Message Header";
        VirginAPIIntegration: Codeunit "Virgin API Integration";
        Status: Option "To be Processed",Failed,Processed,"Closed Manually","Skip Processing";
        EntryNo: BigInteger;
        AccessToken: Text;
        Response: Text;
        ErrorText: Text;
        JObject: JsonObject;
        JToken: JsonToken;
        HttpContent: HttpContent;
        HttpClient: HttpClient;
        HttpHeaders: HttpHeaders;
        HttpResponse: HttpResponseMessage;
        HttpRequest: HttpRequestMessage;
    begin
        //GenerateAccessToken(AccessToken);
        GenerateAccessToken_1(AccessToken);

        HttpContent.WriteFrom(JsonText);
        HttpContent.GetHeaders(HttpHeaders);
        HttpHeaders.Remove('Content-Type');
        HttpHeaders.Add('Content-Type', 'application/json');
        HttpClient.SetBaseAddress(URL);
        HttpClient.DefaultRequestHeaders.Add('User-Agent', 'Dynamics 365');
        HttpClient.DefaultRequestHeaders().Add('Authorization', 'Bearer ' + AccessToken);
        // Request Message Log
        EntryNo := MessageHeaderL.InsertMessageHeader(MessageHeaderL.Direction::"Outgoing Request", Status::Processed, 0, FreeText, '', JsonText, 'VIRGIN');

        HttpClient.Post(URL, HttpContent, HttpResponse);
        HttpResponse.Content.ReadAs(Response);

        case HttpResponse.HttpStatusCode of
            200 .. 299:
                begin
                    JObject.ReadFrom(Response);
                    JObject.Get('OR_MESSAGE', JToken);
                    if JToken.AsValue().AsText() <> 'No error' then begin
                        ErrorText := JToken.AsValue().AsText();
                        MessageHeaderL.InsertMessageHeader(MessageHeaderL.Direction::"Outgoing Response", Status::Failed, EntryNo, FreeText, ErrorText, Response, 'VIRGIN-Error');
                    end else
                        MessageHeaderL.InsertMessageHeader(MessageHeaderL.Direction::"Outgoing Response", Status::Processed, EntryNo, FreeText, '', Response, 'VIRGIN');
                end;
            400:
                begin
                    JObject.ReadFrom(Response);
                    JObject.Get('Message', JToken);
                    ErrorText := JToken.AsValue().AsText();
                    MessageHeaderL.InsertMessageHeader(MessageHeaderL.Direction::"Outgoing Response", Status::Failed, EntryNo, FreeText, ErrorText, Response, 'VIRGIN-Error');
                end;
            401:
                begin
                    ErrorText := 'UnAuthorised';
                    MessageHeaderL.InsertMessageHeader(MessageHeaderL.Direction::"Outgoing Response", Status::Failed, EntryNo, FreeText, ErrorText, Response, 'VIRGIN-Error');
                end;
        end;

        Commit();
        if ErrorText > '' then
            Error(ErrorText);
    end;

    procedure GenerateAccessToken(var AccessToken: Text)
    var
        OAuth2: Codeunit OAuth2;
        Scopes: List of [Text];
    begin
        AzadeaSetup.TestField("Virgin Client ID");
        AzadeaSetup.TestField("Virgin Client Secret");
        AzadeaSetup.TestField("Virgin Scope"); // https://optimus-uat2.sandbox.operations.dynamics.com/.default
        AzadeaSetup.TestField("Virgin AccessToken URL"); //https://login.microsoftonline.com/82a56f9b-b47c-497f-bd95-cb5715721b67/oauth2/v2.0/token

        Scopes.Add(AzadeaSetup."Virgin Scope");
        OAuth2.AcquireTokenWithClientCredentials(
            AzadeaSetup."Virgin Client ID", AzadeaSetup."Virgin Client Secret", AzadeaSetup."Virgin AccessToken URL", '', Scopes, AccessToken);
        if AccessToken = '' then
            Error(UnableToGenerateAccessTokenErr)
    end;

    procedure GenerateAccessToken_1(var AccessToken: Text)
    var
        HttpClient: HttpClient;
        HttpContent: HttpContent;
        HttpResponseMessage: HttpResponseMessage;
        HttpHeaders: HttpHeaders;
        Url: Text;
        TokenBody: Text;
        ResponseText: Text;
        JObject: JsonObject;
        JToken: JsonToken;
    begin
        AzadeaSetup.TestField("Virgin Client ID");
        AzadeaSetup.TestField("Virgin Client Secret");
        AzadeaSetup.TestField("Virgin Scope"); // https://optimus-uat2.sandbox.operations.dynamics.com/.default
        AzadeaSetup.TestField("Virgin AccessToken URL"); //https://login.microsoftonline.com/82a56f9b-b47c-497f-bd95-cb5715721b67/oauth2/v2.0/token

        TokenBody := BuildTokenBody();
        Url := AzadeaSetup."Virgin AccessToken URL";

        HttpContent.WriteFrom(TokenBody);
        HttpContent.GetHeaders(HttpHeaders);
        HttpHeaders.Clear();
        HttpHeaders.Add('Content-Type', 'application/x-www-form-urlencoded');
        if HttpClient.Post(URL, HttpContent, HttpResponseMessage) then begin
            HttpResponseMessage.Content.ReadAs(ResponseText);
            JObject.ReadFrom(ResponseText);
            if HttpResponseMessage.IsSuccessStatusCode then begin
                JObject.Get('access_token', JToken);
                JToken.WriteTo(AccessToken);
                AccessToken := DelChr(AccessToken, '=', '"');
            end else
                Error(ResponseText);
        end;
        if AccessToken = '' then
            Error(GetLastErrorText());
    end;

    local procedure BuildTokenBody(): Text
    var
        TextBuilderL: TextBuilder;
    begin
        TextBuilderL.AppendLine('grant_type=' + 'client_credentials');
        TextBuilderL.AppendLine('&' + 'client_id=' + AzadeaSetup."Virgin Client ID");
        TextBuilderL.AppendLine('&' + 'client_secret=' + AzadeaSetup."Virgin Client Secret");
        TextBuilderL.Append('&' + 'scope=' + AzadeaSetup."Virgin Scope");
        exit(TextBuilderL.ToText());
    end;

    local procedure SalesInvoiceLinesToJson(DocumentNo: Code[20]; var TotalAmt: Decimal): JsonArray
    var
        SalesInvoiceLine: Record "Sales Invoice Line";
        VirginAPIIntegration: Codeunit "Virgin API Integration";
        BarcodeNo: Code[50];
        JSalesInvLines: JsonArray;
    begin
        TotalAmt := 0;
        if VirginAPIIntegration.FilterSalesInvoiceLineWithVirginBrandCode(SalesInvoiceLine, DocumentNo) then
            repeat
                TotalAmt += SalesInvoiceLine."Amount Including VAT";
                if SalesInvoiceLine."Barcode No." > '' then
                    BarcodeNo := SalesInvoiceLine."Barcode No."
                else
                    BarcodeNo := GetItemBarCode(SalesInvoiceLine."No.", SalesInvoiceLine."Variant Code");

                BuildItemJson(SalesInvoiceLine."Unit Price", SalesInvoiceLine.Quantity, BarcodeNo, JSalesInvLines);
            until SalesInvoiceLine.Next() = 0;

        exit(JSalesInvLines);
    end;

    local procedure SalesCreditMemoLinesToJson(DocumentNo: Code[20]; var TotalAmt: Decimal): JsonArray
    var
        SalesCrMemoLine: Record "Sales Cr.Memo Line";
        VirginAPIIntegration: Codeunit "Virgin API Integration";
        BarcodeNo: Code[50];
        JSalesInvLines: JsonArray;
    begin
        TotalAmt := 0;
        if VirginAPIIntegration.FilterSalesCreditMemoLineWithVirginBrandCode(SalesCrMemoLine, DocumentNo) then
            repeat
                TotalAmt += SalesCrMemoLine."Amount Including VAT";
                if SalesCrMemoLine."Barcode No." > '' then
                    BarcodeNo := SalesCrMemoLine."Barcode No."
                else
                    BarcodeNo := GetItemBarCode(SalesCrMemoLine."No.", SalesCrMemoLine."Variant Code");

                BuildItemJson(SalesCrMemoLine."Unit Price", -SalesCrMemoLine.Quantity, BarcodeNo, JSalesInvLines);
            until SalesCrMemoLine.Next() = 0;

        TotalAmt := TotalAmt * -1;
        exit(JSalesInvLines);
    end;

    local procedure ConvertTimeToSeconds(InputDateTime: DateTime) ReturnTime: Integer
    var
        InputTime: Time;
        TimeIntoText: Text;
        StringList: List of [Text];
        Hours: Integer;
        Minutes: Integer;
        Seconds: Integer;
    begin
        InputTime := DT2Time(InputDateTime);
        TimeIntoText := Format(InputTime, 0, '<Hours24,2>,<Minutes,2>,<Seconds,2>');
        StringList := TimeIntoText.Split(',');
        if Evaluate(Hours, StringList.Get(1)) then
            Hours := Hours * 60 * 60;
        if Evaluate(Minutes, StringList.Get(2)) then
            Minutes := Minutes * 60;
        if Evaluate(Seconds, StringList.Get(3)) then;

        ReturnTime := Hours + Minutes + Seconds;
    end;

    local procedure GetItemBarCode(ItemNo: Code[20]; VariantCode: Code[10]) BarcodeNo: Code[50];
    var
        Barcode: Record Barcode;
        Item: Record Item;
    begin
        BarcodeNo := '';
        if not Item.Get(ItemNo) then
            exit;

        Barcode.SetCurrentKey("Item No.", "Variant Code");
        Barcode.SetRange("Item No.", ItemNo);
        Barcode.SetRange("Variant Code", VariantCode);
        if Barcode.FindFirst() then
            BarcodeNo := Barcode."Barcode No.";
    end;

    local procedure GenerateTRSNumberIfBlank(IntegrationRefrenceNo: Code[100])
    var
        NoSeriesMgt: Codeunit NoSeriesManagement;
    begin
        // Refer same integration refrence number if already linked to document
        if IntegrationRefrenceNo > '' then begin
            TRSNo := IntegrationRefrenceNo;
            exit;
        end;

        AzadeaSetup.TestField("Virgin Integration TRS No.");
        NoSeriesMgt.InitSeries(AzadeaSetup."Virgin Integration TRS No.", '', Today, TRSNo, AzadeaSetup."Virgin Integration TRS No.");
    end;

    local procedure SetJobQueueStatusForInvoice(var SalesInvHdr: Record "Sales Invoice Header"; NewStatus: Option)
    begin
        SalesInvHdr."Virgin API Job Queue Status" := NewStatus;
        if NewStatus = 0 then
            SalesInvHdr."Virgin Sales Integration" := false;
        SalesInvHdr."Integration Refrence No." := TRSNo;
        SalesInvHdr.Modify();
        Commit();
    end;

    local procedure SetJobQueueStatusForReturn(var pSalesCrMemoHeader: Record "Sales Cr.Memo Header"; NewStatus: Option)
    begin
        pSalesCrMemoHeader."Virgin API Job Queue Status" := NewStatus;
        if NewStatus = 0 then
            pSalesCrMemoHeader."Virgin Sales Integration" := false;
        pSalesCrMemoHeader."Integration Refrence No." := TRSNo;
        pSalesCrMemoHeader.Modify();
        Commit();
    end;

    var
        AzadeaSetup: Record "Azadea Setup";
        TRSNo: Code[100];
        UnableToGenerateAccessTokenErr: Label 'Unable to generate access token';
}