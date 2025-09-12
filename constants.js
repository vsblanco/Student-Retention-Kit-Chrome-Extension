// constants.js

// URL to fetch the master list of students. Used in popup.js.
export const MASTER_LIST_URL = "https://prod-10.westus.logic.azure.com:443/workflows/a9e08bd1329c40ffb9bf28bbc35e710a/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=cR_TUW8U-2foOb1XEAPmKxbK-2PLMK_IntYpxd2WOSo";

// URL to trigger a flow when a submission is found. Used in background.js.
export const SUBMISSION_FOUND_URL = "https://prod-12.westus.logic.azure.com:443/workflows/8ac3b733279b4838833c5454b31d005d/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=RccL3_U3q2nYCWoNfmGisH7rBUIyHrx4SD1vc6bzo7w";