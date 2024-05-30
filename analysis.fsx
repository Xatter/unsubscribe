// #r "bin/Debug/net8.0/UnsubscribeMe.dll"

open System.IO
open System.Text.Json

type Header = {
    Name: string
    Value: string
    ETag: string option
}

type Payload = {
    Body: string option
    Filename: string option
    Headers: Header list
    MimeType: string
    PartId: string option
    Parts: string option
    ETag: string option
}

type MessageInfo = {
    HistoryId: int
    Id: string
    InternalDate: int64
    LabelIds: string list
    Payload: Payload
    Raw: string option
    SizeEstimate: int
    Snippet: string
    ThreadId: string
    ETag: string option
}

let f = File.ReadAllLines "allMessages.json"

let messages = f |> Array.map(fun l -> JsonSerializer.Deserialize<MessageInfo>(l))

let (unsubscribable, nounsubscribe) = 
    messages 
    |> Array.map(fun m -> (m.Id, m.Payload.Headers))
    |> Array.partition(fun (_, headers) -> headers |> List.exists (fun h -> h.Name = "List-Unsubscribe"))

let idHeaderPairs = unsubscribable 
                    |> Array.map (fun (id, headers) -> 
                        let unsubscribeHeader = headers |> List.filter(fun h -> h.Name = "List-Unsubscribe") |> List.exactlyOne
                        let fromHeader = headers |> List.filter(fun h -> h.Name = "From")
                        let replyToHeader = headers |> List.filter(fun h -> h.Name = "Reply-To")
                        if (List.isEmpty fromHeader) then (id, unsubscribeHeader.Value, (replyToHeader |> List.exactlyOne).Value)
                        else (id, unsubscribeHeader.Value, (fromHeader |> List.exactlyOne).Value)
                    )

let grouped = idHeaderPairs |> Array.groupBy(fun (_, _, fromHeaders) -> fromHeaders)

printfn "%A" (grouped |> Array.sortBy(fun (group, values) -> values |> Array.length) |> Array.take 50)
