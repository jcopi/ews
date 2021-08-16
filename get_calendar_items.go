package ews

import "encoding/xml"

type Traversal string

const (
    TraversalShallow Traversal = "Shallow"
    TraversalSoftDeleted Traversal = "SoftDeleted"
    TraversalAsscoiated Traversal = "Associated"
)

// https://docs.microsoft.com/en-us/exchange/client-developer/web-service-reference/finditem
type FindItemRequest struct {
    XMLName struct{} `xml:"m:FindItem"`
    Traversal Traversal `xml:"Traversal,attr"`
    ItemShape *ItemShape `xml:"m:ItemShape"`
    //IndexedPageItemView IndexedPageItemView `xml:"m:IndexedPageItemView,omitempty"`
    //FractionalPageItemView FractionalPageItemView `xml:"m:FractionalPageItemView,omitempty"`
    CalendarView CalendarView `xml:"m:CalendarView,omitempty"`
    //ContactsView ContactsView `xml:"m:ContactsView"`
    //GroupBy GroupBy `xml:"m:GroupBy,ompitempty"`
    //DistinguishedGroupBy GroupBy `xml:"m:DistinguishedBroupBy,omitempty"`
    //Restriction Restriction `xml:"m:Restriction,omitempty"`
    //SortOrder SortOrder `xml:"m:SortOrder,omitempty"`
    ParentFolderIds []ParentFolderId `xml:"m:ParentFolderIds"`
    QueryString string `xml:"m:QueryString,omitempty"`
}

type FindItemResponse struct {
    
}

type findItemResponseEnvelope struct {
    XMLName struct{}               `xml:"Envelope"`
	Body    findItemResponseBody   `xml:"Body"`
}

type findItemREsponseBody struct {
    FindItemResponse FindItemResponse `xml:""`   
}
    
// https://docs.microsoft.com/en-us/exchange/client-developer/web-service-reference/itemshape
type ItemShape struct {
    BaseShape BaseShape `xml:"t:BaseShape,omitempty"`
    AdditionalProperties AdditionalProperties `xml:"t:AdditionalProperties,omitempty"`
}

type CalendarView struct {
    MaxEntriesReturned uint `xml:"MaxEntriesReturned,attr,omitempty"`
    StartDate time.Time `xml:"StartDate,attr"`
    EndDate time.Time `xml:"EndDate,attr"`
}

func FindCalendarItems(c Client, r *FindItemRequest) (*FindItemResponse, error) {
    xmlBytes, err := xml.MarshalIndent(r, "", "  ")
    if err != nil {
        return nil, err   
    }
    
    bb, err := c.SendAndReceive(xmlBytes)
    if err != nil {
        return nil, err   
    }
    
    var soapResp findItemResponseEnvelope
    err = xml.Unmarshal(bb, &soapResp)
    if err != nil {
        return nil, err   
    }
    
    if soapResp.Body.FindItemResponse.ResponseClass == ResponseClassError {
        return nil, errors.New(soapResp.Body.FindItemResponse.MessageText)   
    }
    
    return &soapResp.Body.FindItemResponse, nil
}
