import * as React from "react";
import styles from "./BpclWebStories.module.scss";
import "bootstrap/dist/css/bootstrap.min.css";
import { SPFI, spfi } from "@pnp/sp";
import { SPFx } from "@pnp/sp/presets/all";

import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { Button, Form, Modal, Pagination} from "react-bootstrap";
import "bootstrap-icons/font/bootstrap-icons.css";

import FlipBookViewer from "./FlipBookViewer";
import {
  PeoplePicker,
  PrincipalType
} from "@pnp/spfx-controls-react/lib/PeoplePicker";

interface IProps {
  context: any;
  libraryName: string;
}

const BpclWebStories: React.FC<IProps> = ({ context, libraryName }) => {

  const [showShareModal, setShowShareModal] = React.useState(false);
  const [docs, setDocs] = React.useState<any[]>([]);
  const [selectedFile, setSelectedFile] = React.useState<string | null>(null);

  const [selectedItemId, setSelectedItemId] = React.useState<number>();
  const [selectedUsers, setSelectedUsers] = React.useState<any[]>([]);
  const [searchText, setSearchText] = React.useState("");

  const [currentPage, setCurrentPage] = React.useState(1);
const [pageSize, setPageSize] = React.useState(10);

  // ✅ Create SP instance (PnP v3)
  const sp: SPFI = React.useMemo(() => {
    return spfi().using(SPFx(context));
  }, [context]);

  React.useEffect(() => {
    fetchDocuments();
  }, []);


  const fetchDocuments = async () => {
    try {

      const currentUserId = context.pageContext.legacyPageContext.userId;

      const items = await sp.web.lists
        .getByTitle("Corp_DL_WebStories")
        .items
        .select(
          "Id",
          "FileLeafRef",
          "FileRef",
          "Created",
          "AuthorId",
          "Description0"
        )
        .filter(
          `FSObjType eq 0 and substringof('.pdf', FileLeafRef) and AuthorId eq ${currentUserId}`
        )();

      setDocs(items);
    } catch (error) {
      console.error("Error fetching documents:", error);
    }
  };

const filteredDocs = docs.filter((doc) =>
  doc.FileLeafRef?.toLowerCase().includes(searchText.toLowerCase())
);
// pagination logic code
const totalPages = Math.ceil(filteredDocs.length / pageSize);

const startIndex = (currentPage - 1) * pageSize;
const endIndex = startIndex + pageSize;

const pagedDocs = filteredDocs.slice(startIndex, endIndex);

const getDigest = async (): Promise<string> => {
  const response = await fetch(
    `${context.pageContext.web.absoluteUrl}/_api/contextinfo`,
    {
      method: "POST",
      headers: {
        Accept: "application/json;odata=nometadata"
      }
    }
  );

  const data = await response.json();
  return data.FormDigestValue;
};

const handleShare = async () => {
  try {

    if (!selectedItemId) {
      alert("No document selected.");
      return;
    }

    if (selectedUsers.length === 0) {
      alert("Please select at least one user.");
      return;
    }

    // Ensure users and get IDs
    const userIds: number[] = [];
    for (const user of selectedUsers) {
      const ensuredUser = await sp.web.ensureUser(
        user.loginName || user.secondaryText || user.text
      );
      userIds.push(ensuredUser.Id);
    }

    console.log("New User IDs:", userIds);
    // code to get list ListItemEntityTypeFullName
    const list = await sp.web.lists
    .getByTitle("Corp_DL_WebStories")
    .select("ListItemEntityTypeFullName")();

console.log("ListItemEntityTypeFullName :",list.ListItemEntityTypeFullName);

    // Read existing SharedWith users
    const item = await sp.web.lists
      .getByTitle("Corp_DL_WebStories")
      .items
      .getById(selectedItemId)
      .select("SharedWith/Id")
      .expand("SharedWith")();

    const existingIds: number[] =
      item.SharedWith?.map((u: any) => u.Id) || [];

    console.log("Existing IDs:", existingIds);

    // Merge old + new users
    const allUserIds = [...new Set([...existingIds, ...userIds])];

    console.log("All IDs:", allUserIds);

    const digest = await getDigest();

    const payload = {
      __metadata: {
        type: "SP.Data.Corp_x005f_DL_x005f_WebStoriesItem"
      },

      // Keep history "SP.Data.Corp_x005f_DL_x005f_WebStoriesItem"
      SharedWithId: {
        results: allUserIds
      },

      // Only newly selected users
      SendMailToUserId: {
        results: userIds
      }
    };

    // const payload = {
    //   SharedWithId: {
    //     results: allUserIds
    //   },
    //   SendMailToUserId: {
    //     results: userIds
    //   }
    // };

    const response = await fetch(
      `${context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('Corp_DL_WebStories')/items(${selectedItemId})`,
      {
        method: "POST",
        headers: {
          Accept: "application/json;odata=nometadata",
          "Content-Type": "application/json;odata=verbose",
          "X-RequestDigest": digest,
          "IF-MATCH": "*",
          "X-HTTP-Method": "MERGE"
        },
        body: JSON.stringify(payload)
      }
    );

    if (!response.ok) {
      const error = await response.text();
      console.error(error);
      alert("Failed to update document.");
      return;
    }

    alert("Document shared successfully.");

    setSelectedUsers([]);
    setShowShareModal(false);

  } catch (error) {
    console.error(error);
    alert("Something went wrong.");
  }
};

  return (
    <div className={styles.commonSectionstyle}>
      <div className={styles.pageContainer}>
        {/* Banner */}
        <div className={styles.banner}>
          {/* <button className={styles.backIcon}>
            <i className="bi bi-chevron-left" />
          </button> */}
          <div className={styles.bannerIcon}>
            <i className="bi bi-calendar-event-fill" />
          </div>
          <div>
            <h3>Web Stories</h3>
            <p>
              Stay updated with the latest BPCL stories, innovations,
              achievements, initiatives, and organizational highlights.
            </p>
          </div>
        </div>
      </div>
     
      <div className="d-flex flex-column flex-md-row align-items-md-center justify-content-between mt-4 mb-4">
        <h4 className={`${styles.pagetitle} mb-3 mb-md-0`}>
          Total Web Stories  {filteredDocs.length}
        </h4>

        <div className="d-flex align-items-center flex-nowrap">
          <label htmlFor="search" className="me-2 mb-0">
            Search
          </label>

        <input
        id="search"
        type="text"
        className="form-control"
        placeholder="Search your web stories here..."
        style={{ width: "300px" }}
        value={searchText}
       onChange={(e) => {
    setSearchText(e.target.value);
    setCurrentPage(1);
}}
      />
        </div>
      </div>
      {!selectedFile && (
        <div className={styles.cardGrid}>
          {pagedDocs.map((doc, index) => {
            const thumbnailUrl = `${context.pageContext.web.absoluteUrl}/_layouts/15/getpreview.ashx?path=${doc.FileRef}`;

            return (
              <div
                key={index}
                className={styles.card}
                onClick={() =>
                  setSelectedFile(`${window.location.origin}${doc.FileRef}`)
                }
              >
                {/* PDF Preview */}
                <div className={styles.previewContainer}>
                  <img
                    src={thumbnailUrl} // First page thumbnail of PDF
                    alt={doc.FileLeafRef}
                    className={styles.thumbnail}
                  />

                  {/* Share Button */}
               <button
                className={styles.shareBtn}
                onClick={(e) => {
                  e.preventDefault();
                  e.stopPropagation();

                  console.log("Share Clicked");

                  setSelectedItemId(doc.Id);
                  setShowShareModal(true);
                }}
              >
                <i className="bi bi-share-fill"></i>
              </button>
                </div>

                {/* Card Content */}
                <div className={styles.content}>
                  <span className={styles.badge}>Topic</span>

                  <h5 className={styles.title} title={doc.FileLeafRef}>
                    {doc.FileLeafRef}
                  </h5>

                  <div className={styles.date}>
                   {doc.Created
                      ? new Date(doc.Created).toLocaleDateString("en-GB", {
                          day: "2-digit",
                          month: "short",
                          year: "numeric",
                        })
                      : ""}
                  </div>

                <p
                  className={`${styles.description} mb-0`}
                  title={doc.Description || ""}
                >
                  {doc.Description0 || "No description available"}
                </p>
                </div>
              </div>
            );
          })}
        </div>
      )}
      <div className={`${styles.paginationSection} mt-3`}>
        <div className={styles.leftPagination}>
         <span className={styles.pageInfo}>
  Page {currentPage} of {totalPages || 1}
</span>
<Form.Select
    className={styles.pageSize}
    value={pageSize}
    onChange={(e) => {
        setPageSize(Number(e.target.value));
        setCurrentPage(1);
    }}
>
    <option value={10}>10 per page</option>
    <option value={20}>20 per page</option>
    <option value={50}>50 per page</option>
</Form.Select>
        </div>

        {/* <Pagination className="mb-0 justify-content-end flex-wrap">
          <Pagination.First />
          <Pagination.Prev />

          <Pagination.Item active>1</Pagination.Item>
          <Pagination.Item>2</Pagination.Item>
          <Pagination.Item>3</Pagination.Item>

          <Pagination.Next />
          <Pagination.Last />
        </Pagination> */}

        <Pagination className="mb-0 justify-content-end flex-wrap">

    <Pagination.First
        disabled={currentPage === 1}
        onClick={() => setCurrentPage(1)}
    />

    <Pagination.Prev
        disabled={currentPage === 1}
        onClick={() => setCurrentPage(currentPage - 1)}
    />

    {Array.from({ length: totalPages }, (_, i) => i + 1)
        .slice(
            Math.max(0, currentPage - 3),
            Math.min(totalPages, currentPage + 2)
        )
        .map((page) => (
            <Pagination.Item
                key={page}
                active={page === currentPage}
                onClick={() => setCurrentPage(page)}
            >
                {page}
            </Pagination.Item>
        ))}

    <Pagination.Next
        disabled={currentPage === totalPages || totalPages === 0}
        onClick={() => setCurrentPage(currentPage + 1)}
    />

    <Pagination.Last
        disabled={currentPage === totalPages || totalPages === 0}
        onClick={() => setCurrentPage(totalPages)}
    />

</Pagination>
      </div>
      {selectedFile && (
        <FlipBookViewer
          fileUrl={`${selectedFile}`}
          onClose={() => setSelectedFile(null)}
        />
      )}

      <Modal
        show={showShareModal}
       onHide={() => {
            setShowShareModal(false);
            setSelectedUsers([]);
        }}
        centered
      >
        <Modal.Header closeButton>
          <Modal.Title>Web Stories Title</Modal.Title>
        </Modal.Header>
 
        <Modal.Body>
          <Form.Group className="mb-3">
            <Form.Label className={styles.floatingLabelStyle}>
              Select People
            </Form.Label>
 
          <PeoplePicker
            context={context}
            webAbsoluteUrl={context.pageContext.web.absoluteUrl}
            titleText=""
            personSelectionLimit={20}
            ensureUser={true}
            principalTypes={[PrincipalType.User]}
            resolveDelay={500}
            useSubstrateSearch={false}
            onChange={(items) => {
              console.log(items);
              setSelectedUsers(items);
            }}
          />
          </Form.Group>
        </Modal.Body>
 
        <Modal.Footer>
          <Button
            className={styles.cancelBtn}
            onClick={() => {
                setShowShareModal(false);
                setSelectedUsers([]);
            }}
             
          >
            Cancel
          </Button>
 
         <Button
            className={styles.primaryBtn}
            onClick={handleShare}
        >
            Share
        </Button>
        </Modal.Footer>
      </Modal>

    </div>
  );
};

export default BpclWebStories;
