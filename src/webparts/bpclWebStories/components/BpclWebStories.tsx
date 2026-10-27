
import * as React from "react";
import styles from "./BpclWebStories.module.scss";
import "bootstrap/dist/css/bootstrap.min.css";
import { SPFI, spfi } from "@pnp/sp";
import { SPFx } from "@pnp/sp/presets/all";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";

import {
  Alert,
  Button,
  Form,
  Modal,
  Pagination
} from "react-bootstrap";

import "bootstrap-icons/font/bootstrap-icons.css";
import { Row } from "react-bootstrap";
import FlipBookViewer from "./FlipBookViewer";
import {
  PeoplePicker,
  PrincipalType
} from "@pnp/spfx-controls-react/lib/PeoplePicker";
import webStoriesIcon from "../assets/FlipBookIcon.png";
import { MSGraphClientV3 } from "@microsoft/sp-http";

interface IProps {
  context: any;
  libraryName: string;
}


interface IWebStory {
  Id: number;
  FileLeafRef: string;
  FileRef: string;
  Created: string;
  AuthorId: number;
EncodedAbsUrl: string;
  Author?: {
    Title: string;
  };

  Description0?: string;

    ThumbnailUrl?: string;

  SharedWith?: {
    Id: number;
  }[];
}


const BpclWebStories: React.FC<IProps> = ({
  context
  
}) => {

  const [showShareModal, setShowShareModal] =
    React.useState(false);

  const [docs, setDocs] =
    React.useState<IWebStory[]>([]);

  const [selectedFile, setSelectedFile] =
    React.useState<string | null>(null);

  const [selectedItemId, setSelectedItemId] =
    React.useState<number>();

  const [selectedUsers, setSelectedUsers] =
    React.useState<any[]>([]);

  const [searchText, setSearchText] =
    React.useState("");

  // Year filter
  const currentYear = new Date().getFullYear();

  const [selectedYear, setSelectedYear] =
    React.useState(currentYear.toString());

  const [currentPage, setCurrentPage] =
    React.useState(1);

  const [pageSize, setPageSize] =
    React.useState(10);

  const [successMessage, setSuccessMessage] =
    React.useState("");

    const [isLoading, setIsLoading] = React.useState(true);

    const [isLinkCopied, setIsLinkCopied] = React.useState(false);


  // Create SP instance
  const sp: SPFI = React.useMemo(() => {
    return spfi().using(SPFx(context));
  }, [context]);


  // ============================================================
  // FETCH DOCUMENTS
  // ============================================================


const libraryName = "Corp_DL_WebStories";

const getThumbnailUrl = async (
  itemId: number,
  listId: string
): Promise<string> => {
  try {
    const graphClient: MSGraphClientV3 =
      await context.msGraphClientFactory.getClient("3");

    const siteId =
      context.pageContext.site.id.toString();

    const driveItem =
      await graphClient
        .api(
          `/sites/${siteId}/lists/${listId}/items/${itemId}/driveItem`
        )
        .select("id,parentReference,name")
        .get();

    if (
      !driveItem ||
      !driveItem.id ||
      !driveItem.parentReference?.driveId
    ) {
      return "";
    }

    const driveId =
      driveItem.parentReference.driveId;

    const driveItemId =
      driveItem.id;

    const thumbnailResponse =
      await graphClient
        .api(
          `/drives/${driveId}/items/${driveItemId}/thumbnails`
        )
        .select("medium,large")
        .get();

    if (thumbnailResponse?.value?.length) {
      const thumbnail =
        thumbnailResponse.value[0];

      if (thumbnail.large?.url) {
        return thumbnail.large.url;
      }

      if (thumbnail.medium?.url) {
        return thumbnail.medium.url;
      }
    }

    return "";
  } catch (error) {
    console.error(
      "Thumbnail error for item:",
      itemId,
      error
    );

    return "";
  }
};

const fetchDocuments = async (): Promise<void> => {
  setIsLoading(true);

  try {
    // Get List GUID dynamically using library name
    const list = await sp.web.lists
      .getByTitle(libraryName)
      .select("Id")();

    const listId = list.Id;

    console.log("Library Name:", libraryName);
    console.log("Dynamic List ID:", listId);

    // Get PDF documents
    const items = await sp.web.lists
      .getByTitle(libraryName)
      .items
      .select(
        "Id",
        "FileLeafRef",
        "FileRef",
        "EncodedAbsUrl",
        "Created",
        "AuthorId",
        "Author/Title",
        "Description0"
      )
      .expand("Author")
      .filter(
        `FSObjType eq 0 and substringof('.pdf', FileLeafRef)`
      )
      .top(5000)();

    console.log("PDF ITEMS:", items);

    // Get thumbnails
    const documentsWithThumbnails =
      await Promise.all(
        items.map(
          async (item: IWebStory) => {
            const thumbnailUrl =
              await getThumbnailUrl(
                item.Id,
                listId
              );

            return {
              ...item,
              ThumbnailUrl: thumbnailUrl
            };
          }
        )
      );

    console.log(
      "PDF ITEMS WITH THUMBNAILS:",
      documentsWithThumbnails
    );

    setDocs(
      documentsWithThumbnails
    );

  } catch (error) {
    console.error(
      "Error fetching documents:",
      error
    );
  } finally {
    setIsLoading(false);
  }
};


  React.useEffect(() => {

    fetchDocuments().catch((error) => {

      console.error(
        "Error fetching documents:",
        error
      );

    });

  }, []);


  // ============================================================
  // SEARCH + YEAR FILTER
  // ============================================================

  const filteredDocs = docs.filter((doc) => {

    const search =
      searchText.trim().toLowerCase();

    // Year filter
    const docYear = doc.Created
      ? new Date(doc.Created)
          .getFullYear()
          .toString()
      : "";

    const yearMatch =
      docYear === selectedYear;


    // Search filter
    if (!search) {
      return yearMatch;
    }


    const fileName =
      doc.FileLeafRef
        ?.replace(/\.pdf$/i, "")
        .toLowerCase() || "";


    const description =
      doc.Description0?.toLowerCase() || "";


    const searchMatch =
      fileName.includes(search) ||
      description.includes(search);

     // const EncodedAbsUrl = doc.EncodedAbsUrl?.toLowerCase() || "";


    return yearMatch && searchMatch;

  });


  // ============================================================
  // PAGINATION
  // ============================================================

  const totalPages =
    Math.ceil(
      filteredDocs.length / pageSize
    );


  const startIndex =
    (currentPage - 1) * pageSize;


  const endIndex =
    startIndex + pageSize;


  const pagedDocs =
    filteredDocs.slice(
      startIndex,
      endIndex
    );


  // ============================================================
  // GET DIGEST
  // ============================================================

  const getDigest = async (): Promise<string> => {

    const response = await fetch(
      `${context.pageContext.web.absoluteUrl}/_api/contextinfo`,
      {
        method: "POST",

        headers: {
          Accept:
            "application/json;odata=nometadata"
        }
      }
    );

    const data =
      await response.json();

    return data.FormDigestValue;
  };


  // ============================================================
  // SHARE DOCUMENT
  // ============================================================

  const handleShare = async (): Promise<void> => {

    try {

      if (!selectedItemId) {

        alert(
          "No document selected."
        );

        return;
      }


      if (selectedUsers.length === 0) {

        alert(
          "Please select at least one user."
        );

        return;
      }


      // Ensure users and get IDs

      const userIds: number[] = [];


      for (const user of selectedUsers) {

        const ensuredUser =
          await sp.web.ensureUser(
            user.loginName ||
            user.secondaryText ||
            user.text
          );

        userIds.push(
          ensuredUser.Id
        );

      }


      console.log(
        "New User IDs:",
        userIds
      );


      // Get list entity type

      const list =
        await sp.web.lists
          .getByTitle(
            "Corp_DL_WebStories"
          )
          .select(
            "ListItemEntityTypeFullName"
          )();


      console.log(
        "ListItemEntityTypeFullName :",
        list.ListItemEntityTypeFullName
      );


      // Read existing SharedWith users

      const item =
        await sp.web.lists
          .getByTitle(
            "Corp_DL_WebStories"
          )
          .items
          .getById(
            selectedItemId
          )
          .select(
            "SharedWith/Id"
          )
          .expand(
            "SharedWith"
          )();


      const existingIds: number[] =
        item.SharedWith?.map(
          (u: { Id: number }) =>
            u.Id
        ) || [];


      console.log(
        "Existing IDs:",
        existingIds
      );


      // Merge old + new users

      const allUserIds =
        [
          ...new Set([
            ...existingIds,
            ...userIds
          ])
        ];


      console.log(
        "All IDs:",
        allUserIds
      );


      const digest =
        await getDigest();


      const payload = {

        __metadata: {
          type:
            "SP.Data.Corp_x005f_DL_x005f_WebStoriesItem"
        },

        // Keep history

        SharedWithId: {
          results: allUserIds
        },

        // Only newly selected users

        SendMailToUserId: {
          results: userIds
        }

      };


      const response =
        await fetch(
          `${context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('Corp_DL_WebStories')/items(${selectedItemId})`,
          {
            method: "POST",

            headers: {

              Accept:
                "application/json;odata=nometadata",

              "Content-Type":
                "application/json;odata=verbose",

              "X-RequestDigest":
                digest,

              "IF-MATCH":
                "*",

              "X-HTTP-Method":
                "MERGE"

            },

            body:
              JSON.stringify(payload)

          }
        );


      if (!response.ok) {

        const error =
          await response.text();

        console.error(error);

        alert(
          "Failed to update document."
        );

        return;
      }


      setSuccessMessage(
        "Document shared successfully."
      );


      setTimeout(() => {
        setSuccessMessage("");
      }, 3000);

      setSelectedUsers([]);
      setShowShareModal(false);

    } catch (error) {
      console.error(error);
      alert(
        "Something went wrong."
      );

    }

  };
// copy share link

// Copy share link URL dynamically
const shareUrl = selectedItemId
  ? `${context.pageContext.web.absoluteUrl}/SitePages/ViewWebStories.aspx?documentdetails=${selectedItemId}`
  : "";

// Copy URL functionality
const handleCopyLink = async (): Promise<void> => {
  if (!shareUrl) {
    return;
  }

  try {
    await navigator.clipboard.writeText(shareUrl);
    setIsLinkCopied(true);
  } catch (error) {
    console.error("Error copying link:", error);
  }
};
  // ============================================================
  // UI
  // ============================================================

  return (

    <div
      className={
        styles.commonSectionstyle
      }
    >

      {/* Success Message */}

      {successMessage && (

        <Alert
          variant="success"
          dismissible
          onClose={() =>
            setSuccessMessage("")
          }
        >

          {successMessage}

        </Alert>

      )}


      {/* ======================================================
          LIST PAGE
          Everything inside this block disappears when
          FlipBook is opened.
      ======================================================= */}

      {!selectedFile && (

        <>

          {/* ===============================
              Banner
          ================================ */}

          <div
            className={
              styles.pageContainer
            }
          >

            <div
              className={
                styles.banner
              }
            >

              <div
                className={
                  styles.bannerIcon
                }
              >

                <img className={styles.webStoiesImg}
                  src={webStoriesIcon}
                  alt="Web Stories"
                />

              </div>


              <div>

                <h3>
                  Web Stories
                </h3>

                <p>

                  Explore articles, monthly editions
                  and quarterly publications in an

                  {" "}

                  <strong>
                    interactive flipbook experience.
                  </strong>

                </p>

              </div>

            </div>

          </div>


          {/* ===============================
              Search + Year
          ================================ */}

          <div className="px-4 mb-4">

            <div
              className="d-flex flex-column flex-md-row align-items-md-center justify-content-between mt-2 mb-3"
            >

              <h4
                className={`${styles.pagetitle} mb-3 mb-md-0`}
              >

                Total Web Stories{" "}

                {filteredDocs.length}

              </h4>


              <div
                className="d-flex align-items-center flex-nowrap"
              >

                {/* Year Dropdown */}

                <label
                  htmlFor="year"
                  className="me-2 mb-0"
                >
                  Year
                </label>


                <Form.Select
                  id="year"
                  className="me-3"
                  style={{
                    width: "120px"
                  }}
                  value={selectedYear}
                  onChange={(e) => {

                    setSelectedYear(
                      e.target.value
                    );

                    setCurrentPage(1);

                  }}
                >

                  <option
                    value={currentYear}
                  >
                    {currentYear}
                  </option>


                  <option
                    value={currentYear - 1}
                  >
                    {currentYear - 1}
                  </option>


                  <option
                    value={currentYear - 2}
                  >
                    {currentYear - 2}
                  </option>

                </Form.Select>


                {/* Search */}

                <label
                  htmlFor="search"
                  className="me-2 mb-0"
                >
                  Search
                </label>


                <input
                  id="search"
                  type="text"
                  className="form-control"
                  placeholder="Search your web stories here..."
                  style={{
                    width: "300px"
                  }}
                  value={searchText}
                  onChange={(e) => {

                    setSearchText(
                      e.target.value
                    );

                    setCurrentPage(1);

                  }}
                />

              </div>

            </div>


            {/* ===============================
                Web Story Cards
            ================================ */}

            {isLoading ? (
              <div className="text-center py-5">
                <div className="spinner-border" role="status">
                  <span className="visually-hidden">
                    Loading...
                  </span>
                </div>
              </div>
            ) : pagedDocs.length > 0 ? (
              <Row className="g-4">
                {pagedDocs.map(
                  (doc, index) => {
                    // const thumbnailUrl =
                    //   `${context.pageContext.web.absoluteUrl}/_layouts/15/getpreview.ashx?path=${doc.FileRef}`;

return (
                      <div
                        key={index}
                        className="col-md-2"
                      >
                        <div
                          className={
                            styles.card
                          }
                          onClick={() =>
                            setSelectedFile(
                              `${window.location.origin}${doc.FileRef}`
                            )
                          }
                        >

                          {/* PDF Preview */}

                          <div
                            className={
                              styles.previewContainer
                            }
                          >

                          <img
                            src={
                              doc.ThumbnailUrl ||
                              webStoriesIcon
                            }
                            alt={doc.FileLeafRef}
                            className={styles.thumbnail}
                            onError={(e) => {
                              console.log(
                                "Thumbnail failed:",
                                doc.FileLeafRef
                              );

                              e.currentTarget.src =
                                webStoriesIcon;
                            }}
                          />


                            {/* Share Button */}

                            <button
                              className={
                                styles.shareBtn
                              }
                              onClick={(e) => {

                                e.preventDefault();

                                e.stopPropagation();

                                setSelectedItemId(doc.Id);
                                setIsLinkCopied(false); //new added to reset the copy state when opening the modal
                                setShowShareModal(true);

                              }}
                            >

                              <i className="bi bi-share-fill" />

                            </button>

                          </div>


                          {/* Card Content */}

                           <div className={styles.content}>
                           
                          <h5
                            className={styles.title}
                            //   title={
                            //     doc.FileLeafRef?.replace(
                            //       /\.pdf$/i,
                            //       ""
                            //     )
                            //   }
                            // >
 
                            //   {doc.FileLeafRef?.replace(
                            //     /\.pdf$/i,
                            //     ""
                            //   )}<p
                            // className={`${styles.description} mb-0`}
                            title={doc.Description0 || ""}
                          >
                            {doc.Description0 || "No description available"}
 
                            {/* </p> */}
                          </h5>
 
                          {/* Author */}
 
                          <div className={styles.author}>
                            <i className="bi bi-person-fill me-1" />
 
                            {doc.Author?.Title || "Unknown"}
                          </div>
 
                          {/* Date */}
 
                          <div className={styles.date}>
                            {doc.Created
                              ? new Date(doc.Created).toLocaleDateString(
                                  "en-GB",
                                  {
                                    day: "2-digit",
                                    month: "short",
                                    year: "numeric",
                                  },
                                )
                              : ""}
                          </div>
 
                          {/* Description */}
 
                          {/* <p
                              className={`${styles.description} mb-0`}
                              title={
                                doc.Description0 ||
                                ""
                              }
                            >
 
                              {doc.Description0 ||
                                "No description available"}
 
                            </p> */}
                          {/* checking */}
                          {/* <div
                            className={styles.description}
                            title={doc.FileLeafRef?.replace(/\.pdf$/i, "")}
                          >
                            {doc.FileLeafRef?.replace(/\.pdf$/i, "")}
 
                            <span
                              className={`${styles.description} mb-0`}
                            ></span>
                          </div> */}
 
                          {/* checking */}
                        </div>

                        </div>

                      </div>

                    );

                  }
                )}

              </Row>

            ) : (

              /* No Records */

              <div
                className="text-center py-5"
              >

                <h5>
                  No web stories found
                </h5>

                <p
                  className="text-muted mb-0"
                >

                  No records found for the
                  selected year or search criteria.

                </p>

              </div>

            )}


            {/* ===============================
                Pagination
            ================================ */}

            <div
              className={`${styles.paginationSection} mt-3`}
            >

              <div
                className={
                  styles.leftPagination
                }
              >

                <span
                  className={
                    styles.pageInfo
                  }
                >

                  Page {currentPage} of{" "}
                  {totalPages || 1}

                </span>


                <Form.Select
                  className={
                    styles.pageSize
                  }
                  value={pageSize}
                  onChange={(e) => {

                    setPageSize(
                      Number(
                        e.target.value
                      )
                    );

                    setCurrentPage(1);

                  }}
                >

                  <option value={10}>
                    10 per page
                  </option>

                  <option value={20}>
                    20 per page
                  </option>

                  <option value={50}>
                    50 per page
                  </option>

                </Form.Select>

              </div>


              <Pagination
                className="mb-0 justify-content-end flex-wrap"
              >

                <Pagination.First
                  disabled={
                    currentPage === 1
                  }
                  onClick={() =>
                    setCurrentPage(1)
                  }
                />


                <Pagination.Prev
                  disabled={
                    currentPage === 1
                  }
                  onClick={() =>
                    setCurrentPage(
                      currentPage - 1
                    )
                  }
                />


                {Array.from(
                  {
                    length: totalPages
                  },
                  (_, i) => i + 1
                )
                  .slice(
                    Math.max(
                      0,
                      currentPage - 3
                    ),
                    Math.min(
                      totalPages,
                      currentPage + 2
                    )
                  )
                  .map((page) => (

                    <Pagination.Item
                      key={page}
                      active={
                        page === currentPage
                      }
                      onClick={() =>
                        setCurrentPage(
                          page
                        )
                      }
                    >

                      {page}

                    </Pagination.Item>

                  ))}


                <Pagination.Next
                  disabled={
                    currentPage === totalPages ||
                    totalPages === 0
                  }
                  onClick={() =>
                    setCurrentPage(
                      currentPage + 1
                    )
                  }
                />


                <Pagination.Last
                  disabled={
                    currentPage === totalPages ||
                    totalPages === 0
                  }
                  onClick={() =>
                    setCurrentPage(
                      totalPages
                    )
                  }
                />

              </Pagination>

            </div>

          </div>

        </>

      )}


      {/* ======================================================
          FLIPBOOK
          Only FlipBookViewer is shown when selectedFile exists.
      ======================================================= */}

      {selectedFile && (

        <FlipBookViewer
          fileUrl={selectedFile}
          onClose={() =>
            setSelectedFile(null)
          }
        />

      )}


      {/* ======================================================
          SHARE MODAL
          Kept unchanged
      ======================================================= */}

      <Modal
        show={showShareModal}
        onHide={() => {
          setShowShareModal(false);
          setSelectedUsers([]);
        }}
        centered
      >

        <Modal.Header closeButton>
          <Modal.Title>
            Web Stories Sharing
          </Modal.Title>
        </Modal.Header>

        <Modal.Body>
           {/* Copy Link URL */}
    <Form.Group className="mb-3">
      <Form.Label className={styles.floatingLabelStyle}>
        Copy Link URL
      </Form.Label>

      <div className="d-flex">
        <Form.Control
          type="text"
          value={shareUrl}
          readOnly
          placeholder="Select a Web Story to generate the link"
        />

     <Button
          // className={styles.primaryBtn}
          onClick={handleCopyLink}
          disabled={!shareUrl}
          style={{
            marginLeft: "8px",
            backgroundColor: isLinkCopied ? "#198754" : "#0d6efd",
            border: isLinkCopied ? "1px solid #198754" : "1px solid #0d6efd",
            color: "#ffffff"
          }}
        >
          {isLinkCopied ? (
            <>
              <i className="bi bi-check2 me-1" />
              {/* Copied */}
            </>
          ) : (
            "Copy"
          )}
        </Button>
      </div>
    </Form.Group>

    {/* Existing Select People section */}
          <Form.Group  className="mb-3">
            <Form.Label
              className={
                styles.floatingLabelStyle
              }
            >
              Select People
            </Form.Label>
            <PeoplePicker
              context={context}
              webAbsoluteUrl={
                context.pageContext.web.absoluteUrl
              }
              titleText=""
              personSelectionLimit={20}
              ensureUser={true}
              principalTypes={[
                PrincipalType.User
              ]}
              resolveDelay={500}
              useSubstrateSearch={false}
              onChange={(items) => {
                console.log(items);
                setSelectedUsers(
                  items
                );

              }}
            />

          </Form.Group>

        </Modal.Body>


        <Modal.Footer>

          <Button
            className={
              styles.cancelBtn
            }
            onClick={() => {

              setShowShareModal(false);

              setSelectedUsers([]);

            }}
          >
            Cancel
          </Button>


          <Button
            className={
              styles.primaryBtn
            }
            onClick={
              handleShare
            }
          >
            Share
          </Button>

        </Modal.Footer>

      </Modal>

    </div>

  );

};


export default BpclWebStories;

