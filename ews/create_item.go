package ews

import (
	"encoding/xml"
	"time"
)

// https://msdn.microsoft.com/en-us/library/office/aa563009(v=exchg.140).aspx

type CreateItem struct {
	XMLName                struct{}          `xml:"m:CreateItem"`
	MessageDisposition     string            `xml:"MessageDisposition,attr"`
	SendMeetingInvitations string            `xml:"SendMeetingInvitations,attr"`
	SavedItemFolderId      SavedItemFolderId `xml:"m:SavedItemFolderId"`
	Items                  Items             `xml:"m:Items"`
}

type Items struct {
	Message      []Message      `xml:"t:Message"`
	CalendarItem []CalendarItem `xml:"t:CalendarItem"`
}

type SavedItemFolderId struct {
	DistinguishedFolderId DistinguishedFolderId `xml:"t:DistinguishedFolderId"`
}

type DistinguishedFolderId struct {
	Id string `xml:"Id,attr"`
}

type Message struct {
	ItemClass    string     `xml:"t:ItemClass"`
	Subject      string     `xml:"t:Subject"`
	Body         Body       `xml:"t:Body"`
	Sender       OneMailbox `xml:"t:Sender"`
	ToRecipients XMailbox   `xml:"t:ToRecipients"`
}

type CalendarItem struct {
	Subject                    string      `xml:"t:Subject"`
	Body                       Body        `xml:"t:Body"`
	ReminderIsSet              bool        `xml:"t:ReminderIsSet"`
	ReminderMinutesBeforeStart int         `xml:"t:ReminderMinutesBeforeStart"`
	Start                      time.Time   `xml:"t:Start"`
	End                        time.Time   `xml:"t:End"`
	IsAllDayEvent              bool        `xml:"t:IsAllDayEvent"`
	LegacyFreeBusyStatus       string      `xml:"t:LegacyFreeBusyStatus"`
	Location                   string      `xml:"t:Location"`
	RequiredAttendees          []Attendees `xml:"t:RequiredAttendees"`
}

type Body struct {
	BodyType string `xml:"BodyType,attr"`
	Body     []byte `xml:",chardata"`
}

type OneMailbox struct {
	Mailbox Mailbox `xml:"t:Mailbox"`
}

type XMailbox struct {
	Mailbox []Mailbox `xml:"t:Mailbox"`
}

type Mailbox struct {
	EmailAddress string `xml:"t:EmailAddress"`
}

type Attendee struct {
	Mailbox Mailbox `xml:"t:Mailbox"`
}

type Attendees struct {
	Attendee []Attendee `xml:"t:Attendee"`
}

// CreateMessageItem
// https://docs.microsoft.com/en-us/exchange/client-developer/web-service-reference/createitem-operation-email-message
func CreateMessageItem(c *Client, m ...Message) error {

	item := &CreateItem{
		MessageDisposition: "SendAndSaveCopy",
		SavedItemFolderId:  SavedItemFolderId{DistinguishedFolderId{Id: "sentitems"}},
	}
	item.Items.Message = append(item.Items.Message, m...)

	xmlBytes, err := xml.MarshalIndent(item, "", "  ")
	if err != nil {
		return err
	}

	_, err = c.sendAndReceive(xmlBytes)
	if err != nil {
		return err
	}
	return nil
}

// SendEmail helper method to send Message
func SendEmail(c *Client, to []string, subject, body string) error {

	m := Message{
		ItemClass: "IPM.Note",
		Subject:   subject,
		Body: Body{
			BodyType: "Text",
			Body:     []byte(body),
		},
		Sender: OneMailbox{
			Mailbox: Mailbox{
				EmailAddress: c.Username,
			},
		},
	}
	mb := make([]Mailbox, len(to))
	for i, addr := range to {
		mb[i].EmailAddress = addr
	}
	m.ToRecipients.Mailbox = append(m.ToRecipients.Mailbox, mb...)

	return CreateMessageItem(c, m)
}

// CreateCalendarItem
// https://docs.microsoft.com/en-us/exchange/client-developer/web-service-reference/createitem-operation-calendar-item
func CreateCalendarItem(c *Client, ci ...CalendarItem) error {

	item := &CreateItem{
		SendMeetingInvitations: "SendToAllAndSaveCopy",
		SavedItemFolderId:      SavedItemFolderId{DistinguishedFolderId{Id: "calendar"}},
	}
	item.Items.CalendarItem = append(item.Items.CalendarItem, ci...)

	xmlBytes, err := xml.MarshalIndent(item, "", "  ")
	if err != nil {
		return err
	}

	_, err = c.sendAndReceive(xmlBytes)
	if err != nil {
		return err
	}
	return nil
}
