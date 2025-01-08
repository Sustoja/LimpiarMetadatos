EXTENSIONS = ['.docx', '.xlsx', '.pptx', '.pdf']
META_FIELDS = {'.xlsx': ["category", "creator", "description", "keywords", "last_modified_by", "subject", "title"],
               '.docx': ["author", "last_modified_by", "category", "comments", "content_status", "identifier",
                         "keywords", "language", "revision", "subject", "title", "version"],
               '.pptx': ["author", "last_modified_by", "category", "comments", "content_status", "identifier",
                         "keywords", "language", "revision", "subject", "title", "version"],
               '.pdf':  ['Author', 'Creator', 'Keywords', 'Producer', "Subject", "Title"]
               }
DATE_FIELDS = {'.xlsx': ["created", "modified"],
               '.docx': ["created", "modified", "last_printed"],
               '.pptx': ["created", "modified", "last_printed"],
               '.pdf':  ["CreationDate", "ModDate"]
               }