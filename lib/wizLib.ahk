; ahk 用 com 方式操作 wiz 

class Wiz {
    ; 存在一个问题，wiz重启后会报错
    static objApp := ComObjCreate("Wiz.WizExplorerApp")
    static DB_PATH := ""
    static db := ComObjCreate("WizKMCore.WizDatabase")
    static sqliteDB := ComObjCreate("WizKMCore.WizSQLiteDatabase")

    __New() {
        objDatabase := this.objApp.Database
        this.DB_PATH := objDatabase.DatabasePath "index.db"
    }

    getAllDocs(limitNum) {
        this.sqliteDB.Open(this.DB_PATH)

        limit := limitNum ? " LIMIT " limitNum : ""
        rowset := this.sqliteDB.SQLQuery("SELECT DOCUMENT_TITLE, DOCUMENT_GUID FROM WIZ_DOCUMENT"
            . limit . " ORDER BY DT_MODIFIED DESC", "")
        docs := []
        while (!rowset.EOF) {
            docs.Push({ title: rowset.GetFieldValue(0), guid: rowset.GetFieldValue(1) })
            rowset.MoveNext()
        }

        this.mergeTagToDocs(docs, this.sqliteDB)

        this.sqliteDB.Close()
        return docs
    }

    mergeTagToDocs(docs, sqliteDB) {
        tagsMap := {}
        sql = 
        (
            SELECT DOCUMENT_GUID, WIZ_TAG.TAG_NAME
            FROM WIZ_DOCUMENT_TAG
            INNER JOIN WIZ_TAG ON WIZ_DOCUMENT_TAG.TAG_GUID =WIZ_TAG.TAG_GUID
        )
        rowset := sqliteDB.SQLQuery(sql, "")
        while (!rowset.EOF) {
            docGuid := rowset.GetFieldValue(0)
            tagName := rowset.GetFieldValue(1)
            rowset.MoveNext()

            if Not tagsMap[docGuid]
                tagsMap[docGuid] := tagName
            else
                tagsMap[docGuid] .= ", " tagName
        }

        For i, doc in docs
        {
            sTag := tagsMap[doc.guid]
            doc.tagText := sTag
            ; doc.search := doc.title " " sTag
        }
    }

    getCurrentDoc() {
        this.objApp.Window.CurrentDocument
    }

    ; wiz 文件夹
    getFolders() {
        folders := []

        ; 按修改日期排序，但是存在子文件夹，没法直接定位到第一个
        sql = 
        (
            SELECT DISTINCT
                DOCUMENT_LOCATION,
                DOCUMENT_GUID
            FROM
                (SELECT * FROM WIZ_DOCUMENT ORDER BY DT_MODIFIED)
            GROUP BY
                DOCUMENT_LOCATION
        )

        this.sqliteDB.Open(this.DB_PATH)

        rowset := this.sqliteDB.SQLQuery(sql, "")
        While (!rowset.EOF) {
            location := rowset.GetFieldValue(0)
            docGuid := rowset.GetFieldValue(1)

            location := this.fixFolderName(location)

            arr := StrSplit(Trim(location, "/"), "/")
            title := arr[arr.Length()]
            displayTitle := repeat("  ", arr.Length() - 1) . title

            folders.Push({ title: title, displayTitle: displayTitle, location: location
                , search: StrReplace(location, "/", " "), guid: docGuid })
            rowset.MoveNext()
        }

        this.sqliteDB.Close()
        return folders
    }
    fixFolderName(name) {
        out := StrReplace(name, "My Journals", "我的日记")
        out := StrReplace(out, "My Notes", "我的笔记")
        return out
    }
    openFolder(info) {
        documentGUID := info.guid
        databasePath := this.objApp.Database.DatabasePath
        KbGUID := this.objApp.Database.KbGUID

        params := "/DatabasePath="  databasePath  " /KbGUID="  KbGUID  " /DocumentGUID="  documentGUID
        this.objApp.Window.ExecCommand("locatedocument", params)
    }
    ; （已弃用）由于获取第一个文档耗时太久，耗时 4.5s
    getFolders2() {
        this.db.Open("")

        folders := []
        this.getFolderInfo(this.db, 0, folders)

        this.db.Close()
        return folders
    }
    getFolderInfo(ByRef wizFolder, level, ByRef arr) {
        try name := wizFolder.Name
        if (name) {
            displayTitle := repeat("  ", level) . name
            childDocs := wizFolder.Documents
            if (childDocs.Count)
                firstDocGuid := childDocs.Item(childDocs.Count- 1).GUID
            else
                firstDocGuid := ""
            info := { "title": name, "displayTitle": displayTitle, "level": level
                , "firstDocGuid": firstDocGuid }

            arr.Push(info)
        }
        
        ; 获取子目录
        folders := wizFolder.Folders
        loop % folders.Count
        {
            i := A_Index - 1
            arr.Push(this.getFolderInfo(folders.Item(i), level + 1, arr))
        }
    }

    getAllTags() {
        sql =
        (
            SELECT WIZ_TAG.TAG_NAME AS name, Count(WIZ_TAG.TAG_GUID) AS count, WIZ_TAG.TAG_GUID
            FROM WIZ_DOCUMENT_TAG
            INNER JOIN WIZ_TAG ON WIZ_DOCUMENT_TAG.TAG_GUID = WIZ_TAG.TAG_GUID
            GROUP BY WIZ_TAG.TAG_NAME
            ORDER BY Count(WIZ_TAG.TAG_GUID) DESC
        )

        this.sqliteDB.Open(this.DB_PATH)

        rowset := this.sqliteDB.SQLQuery(sql, "")
        tags := []
        while (!rowset.EOF) {
            tag := {}
            tag.title := rowset.GetFieldValue(0)
            tag.guid := rowset.GetFieldValue(2)
            tag.displayTitle := tag.title " (" rowset.GetFieldValue(1) ")"
            tags.Push(tag)
            rowset.MoveNext()
        }

        this.sqliteDB.Close()
        return tags
    }
    getAllTags2() {
        this.db.Open("")

        tags := []
        tagObj := this.db.Tags
        loop % tagObj.Count
        {
            i := A_Index - 1
            tag := tagObj.Item(i)
            tags.Push({ title: tag.Name, guid: tag.GUID })
        }
        this.db.Close()

        return tags
    }

    openDoc(docGuid) {
        wizDoc := this.getDocument(docGuid)
        if (wizDoc) {
            this.objApp.Window.ViewDocument(wizDoc, true)
        }
    }

    openTag(info) {
        guid := info.guid
        ; 通过tag guid获得列出对应的文档
        try {
            objDatabase := this.objApp.Database
            objTag := objDatabase.TagFromGUID(guid)
            objTags := this.objApp.CreateWizObject("WizKMCore.WizTagCollection")
            objTags.Add(objTag)
            documents := objDatabase.DocumentsFromTags(objTags)
            this.objApp.Window.DocumentsCtrl.SetDocuments(documents)
        } catch e {
            msgbox "openTag Error"
        }
    }

    sendDocuments(docs) {
        docsFiltered := this.objApp.CreateWizObject("WizKMCore.WizDocumentCollection")
        for i, doc in docs
        {
            wizDoc := this.getDocument(doc.guid)
            if (wizDoc)
                docsFiltered.Add(wizDoc)
        }

        this.objApp.Window.DocumentsCtrl.SetDocuments(docsFiltered)
    }

    newDocument() {
        ;wizCommonUI := ComObjCreate("WizKMControls.WizCommonUI")
        ;wizCommonUI.NewDocument(objApp, [in] IDispatch* pEvents, [in] IDispatch* pFolderDisp, [in] BSTR bstrOptions);
    }
    
    getDocument(docGuid) {
        wizDoc := ""
        
        try {
            objDatabase := this.objApp.Database
            wizDoc := objDatabase.DocumentFromGUID(docGuid)
        } catch e {

        }
        
        return wizDoc
    }

    ; 中间栏 DocumentsCtrl
    getSelectedDocuments() {
        wizDocs := this.objApp.Window.DocumentsCtrl.SelectedDocuments
        ; list all tags
        Loop % wizDocs.Count
        {
            i := A_Index - 1
            doc := wizDocs.Item(i)
            doc.Tags

            ; remove tag
            ; doc.RemoveTag(tag)
        }
    }

}