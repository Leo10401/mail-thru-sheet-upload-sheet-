"use client"

import { useState, useEffect, useCallback, useRef } from "react"
import {
  Search,
  Mail,
  Users,
  FileSpreadsheet,
  Download,
  CheckCircle,
  AlertCircle,
  Copy,
  RefreshCw,
  Moon,
  Sun,
  Upload,
  Link,
  Trash2,
  File,
} from "lucide-react"
import EmailEditor from "react-email-editor"

const API_BASE_URL = "http://localhost:3000"

const XlsxUploadFrontend = () => {
  const [activeTab, setActiveTab] = useState("upload")
  const [loading, setLoading] = useState(false)
  const [data, setData] = useState(null)
  const [error, setError] = useState("")
  const [uploadedFiles, setUploadedFiles] = useState([])
  const [selectedFile, setSelectedFile] = useState(null)
  const [selectedSheet, setSelectedSheet] = useState("")
  const [searchQuery, setSearchQuery] = useState("")
  const [emailSubject, setEmailSubject] = useState("")
  const [emailBody, setEmailBody] = useState("")
  const [emailSendResult, setEmailSendResult] = useState(null)
  const [selectedTemplate, setSelectedTemplate] = useState("")
  const [urlInput, setUrlInput] = useState("")
  const [darkMode, setDarkMode] = useState(() => {
    if (typeof window !== "undefined") {
      return localStorage.getItem("darkMode") === "true" || window.matchMedia("(prefers-color-scheme: dark)").matches
    }
    return false
  })
  const emailEditorRef = useRef(null)
  const [_editorLoaded, setEditorLoaded] = useState(false)
  const [templates, setTemplates] = useState({})
  const fileInputRef = useRef(null)

  // Toggle dark mode
  const toggleDarkMode = () => {
    const newMode = !darkMode
    setDarkMode(newMode)
    if (typeof window !== "undefined") {
      localStorage.setItem("darkMode", newMode)
    }
  }

  // Apply dark mode class to body
  useEffect(() => {
    if (darkMode) {
      document.body.classList.add("dark")
    } else {
      document.body.classList.remove("dark")
    }
  }, [darkMode])

  // Load template into editor
  const loadTemplateInEditor = useCallback(
    async (templateName) => {
      console.log("Loading template:", templateName);
      if (!emailEditorRef.current?.editor || !_editorLoaded) {
        console.error("Editor ref not available or not loaded yet");
        setError("Email editor is not initialized. Please refresh the page.");
        return;
      }

      try {
        let template = templates[templateName];

        // Load the template from server if not already loaded
        if (!template && templateName) {
          console.log(`Loading template from public/templates/${templateName}.json...`);
          const response = await fetch(`/templates/${templateName}.json`);
          if (!response.ok) {
            throw new Error(`HTTP error! status: ${response.status}`);
          }
          template = await response.json();
          setTemplates((prev) => ({ ...prev, [templateName]: template }));
        }

        if (!template) {
          console.error("Template not found or invalid:", templateName);
          setError(`Template '${templateName}' not found or invalid`);
          return;
        }

        console.log("Loaded template object:", template);

        // Validate template structure before attempting to load
        if (!template.body || typeof template.body !== "object") {
          console.error("Invalid template format: missing or invalid 'body'");
          setError("Template format is invalid. Missing 'body' property.");
          return;
        }

        // The editor handles various schema versions. No need to hardcode a specific one.
        // We will ensure that the counters object contains all necessary properties from the editor's perspective.
        const defaultCounters = {
          u_column: 0,
          u_row: 0,
          u_content_text: 0,
          u_content_button: 0,
          u_content_image: 0,
          u_content_divider: 0,
          u_content_html: 0,
          u_content_menu: 0,
          u_content_social: 0,
          u_content_spacing: 0,
          u_content_video: 0
        };

        if (!template.counters) {
          template.counters = {};
        }

        // Merge default counters with existing template counters
        template.counters = { ...defaultCounters, ...template.counters };

        // Create a simpler template structure
        const simpleTemplate = {
          body: {
            id: "body",
            rows: [
              {
                id: "row-1",
                cells: [
                  {
                    id: "cell-1",
                    contents: [
                      {
                        id: "text-1",
                        type: "text",
                        values: {
                          containerPadding: "10px",
                          anchor: "",
                          fontSize: "14px",
                          textAlign: "center",
                          lineHeight: "140%",
                          linkStyle: {
                            inherit: true,
                            linkColor: "#0000ee",
                            linkHoverColor: "#0000ee",
                            linkUnderline: true,
                            linkHoverUnderline: true,
                          },
                          _meta: {
                            htmlID: "u_content_text_1",
                            htmlClassNames: "u_content_text",
                          },
                          selectable: true,
                          removable: true,
                          draggable: true,
                          duplicatable: true,
                          hideable: true,
                          text: '<h1 style="margin: 0px; line-height: 140%; text-align: center; word-wrap: break-word; font-weight: normal; font-family: arial,helvetica,sans-serif; font-size: 28px;"><strong>Congratulations!</strong></h1><p style="font-size: 14px; line-height: 140%; text-align: center; word-wrap: break-word; font-family: arial,helvetica,sans-serif; margin: 0px;"><span style="font-size: 18px; line-height: 25.2px;">We are pleased to inform you that you have successfully completed the course.</span></p>'
                        }
                      }
                    ],
                    values: {
                      _meta: {
                        htmlID: "u_column_1",
                        htmlClassNames: "u_column"
                      },
                      border: {},
                      padding: "0px",
                      backgroundColor: "",
                      selectable: true,
                      draggable: true,
                      removable: true,
                      duplicatable: true,
                      hideable: true
                    }
                  }
                ],
                values: {
                  _meta: {
                    htmlID: "u_row_1",
                    htmlClassNames: "u_row"
                  },
                  backgroundColor: "",
                  backgroundImage: {
                    url: "",
                    fullWidth: true,
                    repeat: "no-repeat",
                    size: "custom",
                    position: "center"
                  },
                  padding: "0px",
                  selectable: true,
                  draggable: true,
                  removable: true,
                  duplicatable: true,
                  hideable: true
                }
              }
            ],
            values: {
              backgroundColor: "#ffffff",
              backgroundImage: {
                url: "",
                fullWidth: true,
                repeat: "no-repeat",
                size: "custom",
                position: "center"
              },
              contentWidth: "600px",
              contentAlign: "center",
              fontFamily: {
                label: "Arial",
                value: "arial,helvetica,sans-serif"
              },
              preheaderText: "",
              linkStyle: {
                body: true,
                linkColor: "#0000ee",
                linkHoverColor: "#0000ee",
                linkUnderline: true,
                linkHoverUnderline: true
              },
              _meta: {
                htmlID: "u_body",
                htmlClassNames: "u_body"
              }
            }
          },
          schemaVersion: 7,
          categories: [],
          counters: {
            u_column: 1,
            u_row: 1,
            u_content_text: 1,
            u_content_button: 0,
            u_content_image: 0,
            u_content_divider: 0,
            u_content_html: 0,
            u_content_menu: 0,
            u_content_social: 0,
            u_content_spacing: 0,
            u_content_video: 0
          },
          metadata: {
            title: "Certificate Template",
            tags: []
          }
        };

        console.log("Attempting to load design into editor...");
        console.log("Template object being loaded:", JSON.stringify(simpleTemplate, null, 2));
        emailEditorRef.current.editor.loadDesign(simpleTemplate);
        setSelectedTemplate(templateName);
        console.log("Design loaded successfully");

      } catch (error) {
        console.error("Error in loadTemplateInEditor:", error);
        setError(`Failed to load template: ${error.message}`);
      }
    },
    [_editorLoaded, templates] // Add templates to dependency array
  );

  // Handle editor load (unlayer instance is ready)
  const handleEditorLoad = useCallback(
    (unlayer) => {
      console.log("Editor loaded!", unlayer);
      setEditorLoaded(true);
    },
    []
  );

  useEffect(() => {
    if (_editorLoaded && selectedTemplate) {
      console.log("Editor loaded and template selected, attempting to load template:", selectedTemplate);
      loadTemplateInEditor(selectedTemplate);
    } else if (_editorLoaded && activeTab === "sendmail" && !selectedTemplate) {
      // This handles the case where the editor is loaded and we are on the sendmail tab,
      // but no template is selected yet. We might want to load a default empty design here.
      console.log("Editor loaded, sendmail tab active, no template selected. Loading empty design.");
      emailEditorRef.current?.editor.loadDesign({
        "body": {
          "id": "jDNO9w6qCq",
          "rows": [{
              "id": "LJhl92Wa_0",
              "cells": [1],
              "columns": [{
                  "id": "bYgy4VLjNW",
                  "contents": [{
                      "id": "aW3sfX4Bff",
                      "type": "html",
                      "values": {
                          "html": "<!DOCTYPE html>\n<html lang=\"en\">\n<head>\n    <meta charset=\"UTF-8\">\n    <meta name=\"viewport\" content=\"width=device-width, initial-scale=1.0\">\n    <title>DigiPodium - Course Completion Certificate</title>\n</head>\n<body style=\"margin: 0; padding: 0; font-family: 'Open Sans', 'Helvetica Neue', Helvetica, Arial, sans-serif; line-height: 1.6; color: #ffffff; background-color: #121b2c; -webkit-text-size-adjust: 100%; -ms-text-size-adjust: 100%; min-width: 100%; width: 100%;\">\n    \n    <!-- Main Container -->\n    <table role=\"presentation\" cellspacing=\"0\" cellpadding=\"0\" border=\"0\" style=\"width: 100%; margin: 0; padding: 0; background-color: #121b2c; border-collapse: collapse; mso-table-lspace: 0pt; mso-table-rspace: 0pt;\">\n        <tr>\n            <td style=\"padding: 0;\">\n                <table role=\"presentation\" cellspacing=\"0\" cellpadding=\"0\" border=\"0\" style=\"width: 100%; max-width: 680px; margin: 0 auto; background-color: #121b2c; border-collapse: collapse; mso-table-lspace: 0pt; mso-table-rspace: 0pt;\">\n                    \n                    <!-- Header with Logo -->\n                    <tr>\n                        <td style=\"text-align: center; padding: 30px 20px;\">\n                            <a href=\"https://www.digipodium.com\" style=\"text-decoration: none;\">\n                                <img src=\"https://raw.githubusercontent.com/digipodium/Email-Templates/master/Images/logo.png\" alt=\"DigiPodium Logo\" style=\"border: 0; height: auto; line-height: 100%; outline: none; text-decoration: none; max-width: 80px; width: 60px; display: block; margin: 0 auto;\">\n                            </a>\n                        </td>\n                    </tr>\n                    \n                    <!-- Hero Section -->\n                    <tr>\n                        <td style=\"text-align: center; padding: 10px;\">\n                            <img src=\"https://raw.githubusercontent.com/digipodium/Email-Templates/master/Images/congrats.png\" alt=\"Congratulations!\" style=\"border: 0; height: auto; line-height: 100%; outline: none; text-decoration: none; max-width: 400px; width: 90%; display: block; margin: 0 auto 20px;\">\n                            \n                            <!-- Certificate Message Box -->\n                            <table role=\"presentation\" cellspacing=\"0\" cellpadding=\"0\" border=\"0\" style=\"width: 100%; background-color: #1a2332; border-radius: 12px; margin: 20px 0; border-left: 4px solid #ffcc00; border-collapse: collapse; mso-table-lspace: 0pt; mso-table-rspace: 0pt;\">\n                                <tr>\n                                    <td style=\"padding: 20px 15px;\">\n                                        <div style=\"color: #ffff00; font-size: 14px; font-weight: bold; margin-bottom: 15px; line-height: 1.5;\">\n                                            To Retain your Certificate, please download it and then save it on your Google Drive, so that it is safe with you for keeps!<br><br>\n                                            You also need to post it on your LinkedIn Account so that it is visible in the \"Licenses & Certification\" column, which will help you in your placements.<br><br>\n                                            ALL THE BEST FOR BRIGHT FUTURE.\n                                        </div>\n                                        <div style=\"color: #ffffff; font-size: 16px; font-weight: bold; margin: 15px 0; text-align: center;\">\n                                            Use #prouddigipod #digipodium #signaturetosuccess<br>\n                                            to post your Certificate!\n                                        </div>\n                                    </td>\n                                </tr>\n                            </table>\n                            \n                            <!-- CTA Button -->\n                            <table role=\"presentation\" cellspacing=\"0\" cellpadding=\"0\" border=\"0\" style=\"margin: 20px auto; border-collapse: collapse; mso-table-lspace: 0pt; mso-table-rspace: 0pt;\">\n                                <tr>\n                                    <td style=\"border-radius: 8px; background-color: #d5b70a; text-align: center;\">\n                                        <a href=\"#\" style=\"background-color: #d5b70a; color: #000000; text-decoration: none; padding: 15px 25px; border-radius: 8px; font-size: 16px; font-weight: bold; display: inline-block; font-family: 'Open Sans', 'Helvetica Neue', Helvetica, Arial, sans-serif;\">Get Your Certificate</a>\n                                    </td>\n                                </tr>\n                            </table>\n                        </td>\n                    </tr>\n                    \n                    <!-- Quote Section -->\n                    <tr>\n                        <td style=\"text-align: center; padding: 15px; font-style: italic; color: #f4cccc; font-size: 16px; border-top: 1px solid #bbbbbb; border-bottom: 1px solid #bbbbbb; margin: 30px 0;\">\n                            \"Technology keeps moving forward - catch up & keep up!\"\n                        </td>\n                    </tr>\n                    \n                    <!-- Course Sections Container -->\n                    <tr>\n                        <td style=\"padding: 10px;\">\n                            \n                            <!-- IT Training Section -->\n                            <table role=\"presentation\" cellspacing=\"0\" cellpadding=\"0\" border=\"0\" style=\"width: 100%; background-color: #1a2332; border-radius: 12px; margin: 20px 0; border-collapse: collapse; mso-table-lspace: 0pt; mso-table-rspace: 0pt;\">\n                                <!-- Mobile: Stack image on top -->\n                                <tr>\n                                    <td style=\"width: 100%; text-align: center; padding: 20px 15px 10px; vertical-align: middle;\">\n                                        <a href=\"https://www.digipodium.com/training.php\" style=\"text-decoration: none;\">\n                                            <img src=\"https://raw.githubusercontent.com/digipodium/Email-Templates/master/Images/it.png\" alt=\"IT Training\" style=\"border: 0; height: auto; line-height: 100%; outline: none; text-decoration: none; max-width: 150px; width: 80%; border-radius: 8px; display: block; margin: 0 auto;\">\n                                        </a>\n                                    </td>\n                                </tr>\n                                <tr>\n                                    <td style=\"width: 100%; padding: 10px 15px 20px; vertical-align: top;\">\n                                        <h2 style=\"color: #369dba; font-size: 20px; font-weight: bold; margin: 0 0 15px 0; font-family: 'Open Sans', 'Helvetica Neue', Helvetica, Arial, sans-serif; text-align: center;\">More Enhancements</h2>\n                                        <table role=\"presentation\" cellspacing=\"0\" cellpadding=\"0\" border=\"0\" style=\"width: 100%; margin: 15px 0; border-collapse: collapse; mso-table-lspace: 0pt; mso-table-rspace: 0pt;\">\n                                            <tr><td style=\"padding: 6px 0; padding-left: 15px; position: relative; color: #ffffff; font-size: 14px;\">✓ Machine Learning</td></tr>\n                                            <tr><td style=\"padding: 6px 0; padding-left: 15px; position: relative; color: #ffffff; font-size: 14px;\">✓ Advanced Machine Learning</td></tr>\n                                            <tr><td style=\"padding: 6px 0; padding-left: 15px; position: relative; color: #ffffff; font-size: 14px;\">✓ Web Designing & Development</td></tr>\n                                            <tr><td style=\"padding: 6px 0; padding-left: 15px; position: relative; color: #ffffff; font-size: 14px;\">✓ Java Programming</td></tr>\n                                            <tr><td style=\"padding: 6px 0; padding-left: 15px; position: relative; color: #ffffff; font-size: 14px;\">✓ Data Structures</td></tr>\n                                            <tr><td style=\"padding: 6px 0; padding-left: 15px; position: relative; color: #ffffff; font-size: 14px;\">✓ Major Project Training</td></tr>\n                                        </table>\n                                        <table role=\"presentation\" cellspacing=\"0\" cellpadding=\"0\" border=\"0\" style=\"margin-top: 15px; border-collapse: collapse; mso-table-lspace: 0pt; mso-table-rspace: 0pt; width: 100%;\">\n                                            <tr>\n                                                <td style=\"text-align: center;\">\n                                                    <table role=\"presentation\" cellspacing=\"0\" cellpadding=\"0\" border=\"0\" style=\"margin: 0 auto; border-collapse: collapse;\">\n                                                        <tr>\n                                                            <td style=\"border-radius: 6px; background-color: #369dba;\">\n                                                                <a href=\"https://www.digipodium.com/training.php\" style=\"background-color: #369dba; color: #ffffff; text-decoration: none; padding: 10px 20px; border-radius: 6px; font-size: 13px; display: inline-block; font-family: 'Open Sans', 'Helvetica Neue', Helvetica, Arial, sans-serif;\">Learn More</a>\n                                                            </td>\n                                                        </tr>\n                                                    </table>\n                                                </td>\n                                            </tr>\n                                        </table>\n                                    </td>\n                                </tr>\n                            </table>\n                            \n                            <!-- Data Analytics Section -->\n                            <table role=\"presentation\" cellspacing=\"0\" cellpadding=\"0\" border=\"0\" style=\"width: 100%; background-color: #1a2332; border-radius: 12px; margin: 20px 0; border-collapse: collapse; mso-table-lspace: 0pt; mso-table-rspace: 0pt;\">\n                                <!-- Mobile: Stack image on top -->\n                                <tr>\n                                    <td style=\"width: 100%; text-align: center; padding: 20px 15px 10px; vertical-align: middle;\">\n                                        <a href=\"https://www.digipodium.com/training.php#DataAnalytics\" style=\"text-decoration: none;\">\n                                            <img src=\"https://raw.githubusercontent.com/digipodium/Email-Templates/master/Images/da.png\" alt=\"Data Analytics\" style=\"border: 0; height: auto; line-height: 100%; outline: none; text-decoration: none; max-width: 150px; width: 80%; border-radius: 8px; display: block; margin: 0 auto;\">\n                                        </a>\n                                    </td>\n                                </tr>\n                                <tr>\n                                    <td style=\"width: 100%; padding: 10px 15px 20px; vertical-align: top;\">\n                                        <h2 style=\"color: #369dba; font-size: 20px; font-weight: bold; margin: 0 0 15px 0; font-family: 'Open Sans', 'Helvetica Neue', Helvetica, Arial, sans-serif; text-align: center;\">Explore More Areas</h2>\n                                        <table role=\"presentation\" cellspacing=\"0\" cellpadding=\"0\" border=\"0\" style=\"width: 100%; margin: 15px 0; border-collapse: collapse; mso-table-lspace: 0pt; mso-table-rspace: 0pt;\">\n                                            <tr><td style=\"padding: 6px 0; padding-left: 15px; position: relative; color: #ffffff; font-size: 14px;\">✓ Advanced Excel</td></tr>\n                                            <tr><td style=\"padding: 6px 0; padding-left: 15px; position: relative; color: #ffffff; font-size: 14px;\">✓ Tableau</td></tr>\n                                            <tr><td style=\"padding: 6px 0; padding-left: 15px; position: relative; color: #ffffff; font-size: 14px;\">✓ Data Analytics Python</td></tr>\n                                            <tr><td style=\"padding: 6px 0; padding-left: 15px; position: relative; color: #ffffff; font-size: 14px;\">✓ Google Analytics</td></tr>\n                                            <tr><td style=\"padding: 6px 0; padding-left: 15px; position: relative; color: #ffffff; font-size: 14px;\">✓ Real Time Working</td></tr>\n                                        </table>\n                                        <table role=\"presentation\" cellspacing=\"0\" cellpadding=\"0\" border=\"0\" style=\"margin-top: 15px; border-collapse: collapse; mso-table-lspace: 0pt; mso-table-rspace: 0pt; width: 100%;\">\n                                            <tr>\n                                                <td style=\"text-align: center;\">\n                                                    <table role=\"presentation\" cellspacing=\"0\" cellpadding=\"0\" border=\"0\" style=\"margin: 0 auto; border-collapse: collapse;\">\n                                                        <tr>\n                                                            <td style=\"border-radius: 6px; background-color: #369dba;\">\n                                                                <a href=\"https://www.digipodium.com/training.php#DataAnalytics\" style=\"background-color: #369dba; color: #ffffff; text-decoration: none; padding: 10px 20px; border-radius: 6px; font-size: 13px; display: inline-block; font-family: 'Open Sans', 'Helvetica Neue', Helvetica, Arial, sans-serif;\">Learn More</a>\n                                                            </td>\n                                                        </tr>\n                                                    </table>\n                                                </td>\n                                            </tr>\n                                        </table>\n                                    </td>\n                                </tr>\n                            </table>\n                            \n                            <!-- Digital Marketing Section -->\n                            <table role=\"presentation\" cellspacing=\"0\" cellpadding=\"0\" border=\"0\" style=\"width: 100%; background-color: #1a2332; border-radius: 12px; margin: 20px 0; border-collapse: collapse; mso-table-lspace: 0pt; mso-table-rspace: 0pt;\">\n                                <!-- Mobile: Stack image on top -->\n                                <tr>\n                                    <td style=\"width: 100%; text-align: center; padding: 20px 15px 10px; vertical-align: middle;\">\n                                        <a href=\"https://www.digipodium.com/dm.php\" style=\"text-decoration: none;\">\n                                            <img src=\"https://raw.githubusercontent.com/digipodium/Email-Templates/master/Images/dm.png\" alt=\"Digital Marketing\" style=\"border: 0; height: auto; line-height: 100%; outline: none; text-decoration: none; max-width: 150px; width: 80%; border-radius: 8px; display: block; margin: 0 auto;\">\n                                        </a>\n                                    </td>\n                                </tr>\n                                <tr>\n                                    <td style=\"width: 100%; padding: 10px 15px 20px; vertical-align: top;\">\n                                        <h2 style=\"color: #369dba; font-size: 20px; font-weight: bold; margin: 0 0 15px 0; font-family: 'Open Sans', 'Helvetica Neue', Helvetica, Arial, sans-serif; text-align: center;\">Digital Marketing Training</h2>\n                                        <table role=\"presentation\" cellspacing=\"0\" cellpadding=\"0\" border=\"0\" style=\"width: 100%; margin: 15px 0; border-collapse: collapse; mso-table-lspace: 0pt; mso-table-rspace: 0pt;\">\n                                            <tr><td style=\"padding: 6px 0; padding-left: 15px; position: relative; color: #ffffff; font-size: 14px;\">✓ Digital Advertising</td></tr>\n                                            <tr><td style=\"padding: 6px 0; padding-left: 15px; position: relative; color: #ffffff; font-size: 14px;\">✓ Social Media Marketing</td></tr>\n                                            <tr><td style=\"padding: 6px 0; padding-left: 15px; position: relative; color: #ffffff; font-size: 14px;\">✓ Search Engine Marketing</td></tr>\n                                            <tr><td style=\"padding: 6px 0; padding-left: 15px; position: relative; color: #ffffff; font-size: 14px;\">✓ Content Management</td></tr>\n                                            <tr><td style=\"padding: 6px 0; padding-left: 15px; position: relative; color: #ffffff; font-size: 14px;\">✓ Web Analytics</td></tr>\n                                        </table>\n                                        <table role=\"presentation\" cellspacing=\"0\" cellpadding=\"0\" border=\"0\" style=\"margin-top: 15px; border-collapse: collapse; mso-table-lspace: 0pt; mso-table-rspace: 0pt; width: 100%;\">\n                                            <tr>\n                                                <td style=\"text-align: center;\">\n                                                    <table role=\"presentation\" cellspacing=\"0\" cellpadding=\"0\" border=\"0\" style=\"margin: 0 auto; border-collapse: collapse;\">\n                                                        <tr>\n                                                            <td style=\"border-radius: 6px; background-color: #369dba;\">\n                                                                <a href=\"https://www.digipodium.com/dm.php\" style=\"background-color: #369dba; color: #ffffff; text-decoration: none; padding: 10px 20px; border-radius: 6px; font-size: 13px; display: inline-block; font-family: 'Open Sans', 'Helvetica Neue', Helvetica, Arial, sans-serif;\">Learn More</a>\n                                                            </td>\n                                                        </tr>\n                                                    </table>\n                                                </td>\n                                            </tr>\n                                        </table>\n                                    </td>\n                                </tr>\n                            </table>\n                        </td>\n                    </tr>\n                    \n                    <!-- Thank You Section -->\n                    <tr>\n                        <td style=\"background-color: #171512; text-align: center; padding: 30px 15px; border-radius: 12px; margin: 30px 0;\">\n                            <div style=\"color: #ffcc00; font-size: 16px; margin-bottom: 15px; line-height: 1.4;\">\n                                Thank you for being a part! If you have any queries or concerns, feel free to contact us!\n                            </div>\n                            <div style=\"color: #ff6600; font-size: 24px; font-style: italic; font-weight: bold; line-height: 1.3;\">\n                                \"Signature to success!!\"\n                            </div>\n                        </td>\n                    </tr>\n                    \n                    <!-- Contact Footer -->\n                    <tr>\n                        <td style=\"background-color: #0b111f; padding: 20px 15px;\">\n                            <table role=\"presentation\" cellspacing=\"0\" cellpadding=\"0\" border=\"0\" style=\"width: 100%; max-width: 680px; margin: 0 auto; border-collapse: collapse; mso-table-lspace: 0pt; mso-table-rspace: 0pt;\">\n                                <!-- Mobile: Stack sections vertically -->\n                                <tr>\n                                    <!-- Social Media Section -->\n                                    <td style=\"width: 100%; padding: 0 0 20px 0; vertical-align: top; text-align: center;\">\n                                        <h3 style=\"color: #ffffff; font-size: 16px; font-weight: bold; margin: 0 0 15px 0; font-family: 'Open Sans', 'Helvetica Neue', Helvetica, Arial, sans-serif;\">Social Media</h3>\n                                        <table role=\"presentation\" cellspacing=\"0\" cellpadding=\"0\" border=\"0\" style=\"margin: 10px auto; border-collapse: collapse; mso-table-lspace: 0pt; mso-table-rspace: 0pt;\">\n                                            <tr>\n                                                <td style=\"padding: 0 8px;\">\n                                                    <a href=\"https://www.facebook.com/summertrainingandinternship2022\" style=\"text-decoration: none;\">\n                                                        <img src=\"https://d2fi4ri5dhpqd1.cloudfront.net/public/resources/social-networks-icon-sets/t-only-logo-white/facebook@2x.png\" alt=\"Facebook\" width=\"24\" height=\"24\" style=\"border: 0; height: auto; line-height: 100%; outline: none; text-decoration: none; display: inline-block;\">\n                                                    </a>\n                                                </td>\n                                                <td style=\"padding: 0 8px;\">\n                                                    <a href=\"https://instagram.com/digipodium_official\" style=\"text-decoration: none;\">\n                                                        <img src=\"https://d2fi4ri5dhpqd1.cloudfront.net/public/resources/social-networks-icon-sets/t-only-logo-white/instagram@2x.png\" alt=\"Instagram\" width=\"24\" height=\"24\" style=\"border: 0; height: auto; line-height: 100%; outline: none; text-decoration: none; display: inline-block;\">\n                                                    </a>\n                                                </td>\n                                                <td style=\"padding: 0 8px;\">\n                                                    <a href=\"https://www.linkedin.com/company/summertrainingandinternship2022/\" style=\"text-decoration: none;\">\n                                                        <img src=\"https://d2fi4ri5dhpqd1.cloudfront.net/public/resources/social-networks-icon-sets/t-only-logo-white/linkedin@2x.png\" alt=\"LinkedIn\" width=\"24\" height=\"24\" style=\"border: 0; height: auto; line-height: 100%; outline: none; text-decoration: none; display: inline-block;\">\n                                                    </a>\n                                                </td>\n                                                <td style=\"padding: 0 8px;\">\n                                                    <a href=\"https://www.youtube.com/channel/UCyob7nX8d2i2Ik8vgpI-qvg/\" style=\"text-decoration: none;\">\n                                                        <img src=\"https://d2fi4ri5dhpqd1.cloudfront.net/public/resources/social-networks-icon-sets/t-only-logo-white/youtube@2x.png\" alt=\"YouTube\" width=\"24\" height=\"24\" style=\"border: 0; height: auto; line-height: 100%; outline: none; text-decoration: none; display: inline-block;\">\n                                                    </a>\n                                                </td>\n                                            </tr>\n                                        </table>\n                                    </td>\n                                </tr>\n                                <tr>\n                                    <!-- Address Section -->\n                                    <td style=\"width: 100%; padding: 0; vertical-align: top; text-align: center;\">\n                                        <h3 style=\"color: #ffffff; font-size: 16px; font-weight: bold; margin: 0 0 15px 0; font-family: 'Open Sans', 'Helvetica Neue', Helvetica, Arial, sans-serif;\">Where to Find Us</h3>\n                                        <div style=\"color: #c0c0c0; font-size: 13px; line-height: 1.4; margin-bottom: 20px; padding: 0 10px;\">\n                                            Lower Ground Floor. Rajaram Kumar Plaza, Behind Moti Mahal Restaurant, Hazratganj, Lucknow - 226001\n                                        </div>\n                                        <table role=\"presentation\" cellspacing=\"0\" cellpadding=\"0\" border=\"0\" style=\"border-collapse: collapse; mso-table-lspace: 0pt; mso-table-rspace: 0pt; margin: 0 auto;\">\n                                            <tr>\n                                                <td style=\"border-radius: 6px; background-color: #e5a715;\">\n                                                    <a href=\"https://www.digipodium.com\" style=\"background-color: #e5a715; color: #ffffff; text-decoration: none; padding: 8px 16px; border-radius: 6px; font-size: 12px; display: inline-block; font-family: 'Open Sans', 'Helvetica Neue', Helvetica, Arial, sans-serif;\">Visit Website</a>\n                                                </td>\n                                            </tr>\n                                        </table>\n                                    </td>\n                                </tr>\n                            </table>\n                        </td>\n                    </tr>\n                    \n                    <!-- Final Logo -->\n                    <tr>\n                        <td style=\"text-align: center; padding: 20px;\">\n                            <img src=\"https://raw.githubusercontent.com/digipodium/Email-Templates/master/Images/logo.png\" alt=\"DigiPodium\" style=\"border: 0; height: auto; line-height: 100%; outline: none; text-decoration: none; max-width: 40px; width: 40px; opacity: 0.8; display: block; margin: 0 auto;\">\n                        </td>\n                    </tr>\n                </table>\n            </td>\n        </tr>\n    </table>\n    \n</body>\n</html>",
                          "hideDesktop": false,
                          "displayCondition": null,
                          "_styleGuide": null,
                          "containerPadding": "10px",
                          "anchor": "",
                          "_meta": {
                              "htmlID": "u_content_html_1",
                              "description": "",
                              "htmlClassNames": "u_content_html"
                          },
                          "selectable": true,
                          "draggable": true,
                          "duplicatable": true,
                          "deletable": true,
                          "hideable": true,
                          "locked": false
                      }
                  }],
                  "values": {
                      "_meta": {
                          "htmlID": "u_column_1",
                          "htmlClassNames": "u_column"
                      },
                      "border": {},
                      "padding": "0px",
                      "deletable": true,
                      "backgroundColor": ""
                  }
              }],
              "values": {
                  "displayCondition": null,
                  "columns": false,
                  "_styleGuide": null,
                  "backgroundColor": "",
                  "columnsBackgroundColor": "",
                  "backgroundImage": {
                      "url": "",
                      "fullWidth": true,
                      "repeat": "no-repeat",
                      "size": "custom",
                      "position": "center"
                  },
                  "padding": "0px",
                  "anchor": "",
                  "hideDesktop": false,
                  "_meta": {
                      "htmlID": "u_row_1",
                      "htmlClassNames": "u_row"
                  },
                  "selectable": true,
                  "draggable": true,
                  "duplicatable": true,
                  "deletable": true,
                  "hideable": true,
                  "locked": false
              }
          }],
          "headers": [],
          "footers": [],
          "values": {
              "_styleGuide": null,
              "popupPosition": "center",
              "popupWidth": "600px",
              "popupHeight": "auto",
              "borderRadius": "10px",
              "contentAlign": "center",
              "contentVerticalAlign": "center",
              "contentWidth": "500px",
              "fontFamily": {
                  "label": "Arial",
                  "value": "arial,helvetica,sans-serif"
              },
              "textColor": "#000000",
              "popupBackgroundColor": "#FFFFFF",
              "popupBackgroundImage": {
                  "url": "",
                  "fullWidth": true,
                  "repeat": "no-repeat",
                  "size": "cover",
                  "position": "center"
              },
              "popupOverlay_backgroundColor": "rgba(0, 0, 0, 0.1)",
              "popupCloseButton_position": "top-right",
              "popupCloseButton_backgroundColor": "#DDDDDD",
              "popupCloseButton_iconColor": "#000000",
              "popupCloseButton_borderRadius": "0px",
              "popupCloseButton_margin": "0px",
              "popupCloseButton_action": {
                  "name": "close_popup",
                  "attrs": {
                      "onClick": "document.querySelector('.u-popup-container').style.display = 'none';"
                  }
              },
              "language": {},
              "backgroundColor": "#F9F9F9",
              "preheaderText": "",
              "linkStyle": {
                  "body": true,
                  "linkColor": "#0000ee",
                  "linkHoverColor": "#0000ee",
                  "linkUnderline": true,
                  "linkHoverUnderline": true
              },
              "backgroundImage": {
                  "url": "",
                  "fullWidth": true,
                  "repeat": "no-repeat",
                  "size": "custom",
                  "position": "center"
              },
              "_meta": {
                  "htmlID": "u_body",
                  "description": "",
                  "htmlClassNames": "u_body"
              }
          }
      },
        "schemaVersion": 7,
        "categories": [],
        "counters": {
          "u_column": 0,
          "u_row": 0,
          "u_content_text": 0,
          "u_content_button": 0,
          "u_content_image": 0,
          "u_content_divider": 0,
          "u_content_html": 0,
          "u_content_menu": 0,
          "u_content_social": 0,
          "u_content_spacing": 0,
          "u_content_video": 0
        },
        "metadata": {
          "title": "Empty Email Template",
          "tags": []
        }
      });
    }
  }, [_editorLoaded, selectedTemplate, loadTemplateInEditor, activeTab]);

  // Handle template selection
  const handleTemplateChange = useCallback(
    (e) => {
      const templateName = e.target.value;
      setSelectedTemplate(templateName);
    },
    []
  );

  // API Functions
  const apiCall = useCallback(async (endpoint, options = {}) => {
    setLoading(true)
    setError("")
    try {
      const response = await fetch(`${API_BASE_URL}${endpoint}`, {
        headers: {
          "Content-Type": "application/json",
        },
        ...options,
      })

      if (!response.ok) {
        throw new Error(`HTTP error! status: ${response.status}`)
      }

      const result = await response.json()
      return result
    } catch (err) {
      setError(err.message)
      throw err
    } finally {
      setLoading(false)
    }
  }, [])

  // File upload functions
  const uploadFile = async (file) => {
    const formData = new FormData()
    formData.append("file", file)

    setLoading(true)
    setError("")
    try {
      const response = await fetch(`${API_BASE_URL}/upload`, {
        method: "POST",
        body: formData,
      })

      if (!response.ok) {
        throw new Error(`Upload failed: ${response.statusText}`)
      }

      const result = await response.json()
      await fetchUploadedFiles()
      setSelectedFile(result.fileId)
      setSelectedSheet(result.sheets[0])
      return result
    } catch (err) {
      setError(err.message)
      throw err
    } finally {
      setLoading(false)
    }
  }

  const uploadFromUrl = async (url) => {
    setLoading(true)
    setError("")
    try {
      const response = await fetch(`${API_BASE_URL}/upload-url`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ url }),
      })
      if (!response.ok) {
        throw new Error(`URL upload failed: ${response.statusText}`)
      }
      const result = await response.json()
      await fetchUploadedFiles()
      setSelectedFile(result.fileId)
      setSelectedSheet(result.sheets[0])
      return result
    } catch (err) {
      setError(err.message)
      throw err
    } finally {
      setLoading(false)
    }
  }

  const fetchUploadedFiles = useCallback(async () => {
    try {
      const result = await apiCall("/files")
      setUploadedFiles(result.files)
    } catch (err) {
      console.error("Failed to fetch uploaded files:", err)
    }
  }, [apiCall])

  const deleteFile = async (fileId) => {
    try {
      await apiCall(`/file/${fileId}`, { method: "DELETE" })
      await fetchUploadedFiles()
      if (selectedFile === fileId) {
        setSelectedFile(null)
        setSelectedSheet("")
        setData(null)
      }
    } catch (err) {
      console.error("Failed to delete file:", err)
    }
  }

  // Data fetching functions
  const fetchFileData = useCallback(async () => {
    if (!selectedFile || !selectedSheet) return
    try {
      const result = await apiCall(`/file/${selectedFile}/${selectedSheet}`)
      setData(result)
    } catch (err) {
      console.error("Failed to fetch file data:", err)
    }
  }, [apiCall, selectedFile, selectedSheet])

  const fetchContacts = useCallback(async () => {
    if (!selectedFile || !selectedSheet) return
    try {
      const result = await apiCall(`/contacts/${selectedFile}/${selectedSheet}`)
      setData(result)
    } catch (err) {
      console.error("Failed to fetch contacts:", err)
    }
  }, [apiCall, selectedFile, selectedSheet])

  const fetchEmails = useCallback(async () => {
    if (!selectedFile || !selectedSheet) return
    try {
      const result = await apiCall(`/emails/${selectedFile}/${selectedSheet}`)
      setData(result)
    } catch (err) {
      console.error("Failed to fetch emails:", err)
    }
  }, [apiCall, selectedFile, selectedSheet])

  const searchContacts = useCallback(async () => {
    if (!searchQuery || !selectedFile || !selectedSheet) return
    try {
      const result = await apiCall(
        `/contacts/${selectedFile}/${selectedSheet}/search?query=${encodeURIComponent(searchQuery)}`,
      )
      setData(result)
    } catch (err) {
      console.error("Failed to search contacts:", err)
    }
  }, [apiCall, searchQuery, selectedFile, selectedSheet])

  // Load uploaded files on mount
  useEffect(() => {
    fetchUploadedFiles()
  }, [fetchUploadedFiles])

  const copyToClipboard = async (text) => {
    try {
      await navigator.clipboard.writeText(text)
    } catch (err) {
      console.error("Failed to copy text:", err)
    }
  }

  const downloadData = () => {
    if (!data) return

    const dataStr = JSON.stringify(data, null, 2)
    const dataUri = "data:application/json;charset=utf-8," + encodeURIComponent(dataStr)
    const exportFileDefaultName = `file-data-${Date.now()}.json`

    const linkElement = document.createElement("a")
    linkElement.setAttribute("href", dataUri)
    linkElement.setAttribute("download", exportFileDefaultName)
    linkElement.click()
  }

  const handleKeyPress = (e) => {
    if (e.key === "Enter" && activeTab === "search") {
      searchContacts()
    }
  }

  const handleFileSelect = (e) => {
    const file = e.target.files[0]
    if (file) {
      uploadFile(file)
    }
  }

  // Send emails
  const sendEmails = async () => {
    if (!selectedFile || !selectedSheet) {
      setError("Please select a file and sheet first")
      return
    }

    setLoading(true)
    setEmailSendResult(null)
    setError("")
    try {
      const htmlContent = await new Promise((resolve) => {
        if (emailEditorRef.current && emailEditorRef.current.editor) {
          emailEditorRef.current.editor.exportHtml((data) => {
            resolve(data.html)
          })
        } else {
          resolve(emailBody)
        }
      })

      const response = await fetch(`${API_BASE_URL}/send-emails/${selectedFile}/${selectedSheet}`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          subject: emailSubject,
          body: htmlContent,
          templateType: "certificate",
        }),
      })
      const result = await response.json()
      if (!response.ok) throw new Error(result.error || "Failed to send emails")
      setEmailSendResult(result)
    } catch (err) {
      setError(err.message)
    } finally {
      setLoading(false)
    }
  }

  return (
    <div
      className={`min-h-screen w-full transition-colors duration-300 ${darkMode ? "bg-gray-900 text-gray-100" : "bg-gradient-to-br from-blue-50 to-indigo-100 text-gray-900"}`}
    >
      <div className="w-auto min-h-screen p-4">
        {/* Header */}
        <div
          className={`${darkMode ? "bg-gray-800 shadow-xl border border-gray-700" : "bg-white shadow-xl"} rounded-2xl p-6 mb-6 w-full transition-colors duration-300`}
        >
          <div className="flex items-center justify-between mb-4">
            <div className="flex items-center gap-3">
              <FileSpreadsheet className={`h-8 w-8 ${darkMode ? "text-purple-400" : "text-indigo-600"}`} />
              <h1 className="text-3xl font-bold">XLSX Upload Dashboard</h1>
            </div>
            <button
              onClick={toggleDarkMode}
              className={`p-2 rounded-full ${darkMode ? "bg-gray-700 hover:bg-gray-600" : "bg-gray-200 hover:bg-gray-300"} transition-colors`}
              aria-label={darkMode ? "Switch to light mode" : "Switch to dark mode"}
            >
              {darkMode ? <Sun className="h-5 w-5 text-yellow-300" /> : <Moon className="h-5 w-5 text-indigo-700" />}
            </button>
          </div>

          {/* File Selection */}
          <div className="space-y-4">
            <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
              <div>
                <label className={`block text-sm font-medium mb-2 ${darkMode ? "text-gray-300" : "text-gray-700"}`}>
                  Selected File
                </label>
                <select
                  value={selectedFile || ""}
                  onChange={(e) => {
                    setSelectedFile(e.target.value)
                    const file = uploadedFiles.find((f) => f.id === e.target.value)
                    if (file && file.sheets.length > 0) {
                      setSelectedSheet(file.sheets[0])
                    }
                  }}
                  className={`px-4 py-3 rounded-lg focus:ring-2 w-full transition-colors ${
                    darkMode
                      ? "bg-gray-700 border-gray-600 focus:ring-purple-500 text-white"
                      : "border border-gray-300 focus:ring-indigo-500 text-gray-900"
                  }`}
                >
                  <option value="">Select a file...</option>
                  {uploadedFiles.map((file) => (
                    <option key={file.id} value={file.id}>
                      {file.name}
                    </option>
                  ))}
                </select>
              </div>

              <div>
                <label className={`block text-sm font-medium mb-2 ${darkMode ? "text-gray-300" : "text-gray-700"}`}>
                  Selected Sheet
                </label>
                <select
                  value={selectedSheet}
                  onChange={(e) => setSelectedSheet(e.target.value)}
                  disabled={!selectedFile}
                  className={`px-4 py-3 rounded-lg focus:ring-2 w-full transition-colors disabled:opacity-50 ${
                    darkMode
                      ? "bg-gray-700 border-gray-600 focus:ring-purple-500 text-white"
                      : "border border-gray-300 focus:ring-indigo-500 text-gray-900"
                  }`}
                >
                  <option value="">Select a sheet...</option>
                  {selectedFile &&
                    uploadedFiles
                      .find((f) => f.id === selectedFile)
                      ?.sheets.map((sheet) => (
                        <option key={sheet} value={sheet}>
                          {sheet}
                        </option>
                      ))}
                </select>
              </div>
            </div>
          </div>
        </div>

        {/* Navigation Tabs */}
        <div
          className={`${darkMode ? "bg-gray-800 border border-gray-700" : "bg-white"} rounded-2xl shadow-xl mb-6 transition-colors duration-300`}
        >
          <div className={`flex overflow-x-auto ${darkMode ? "border-b border-gray-700" : "border-b border-gray-200"}`}>
            {[
              { id: "upload", label: "Upload", icon: Upload },
              { id: "contacts", label: "Contacts", icon: Users },
              { id: "emails", label: "Emails", icon: Mail },
              { id: "sheet", label: "Sheet Data", icon: FileSpreadsheet },
              { id: "search", label: "Search", icon: Search },
              { id: "sendmail", label: "Send Email", icon: Mail },
            ].map((tab) => {
              const Icon = tab.icon
              return (
                <button
                  key={tab.id}
                  onClick={() => setActiveTab(tab.id)}
                  className={`flex items-center gap-2 px-6 py-4 font-medium transition-colors whitespace-nowrap ${
                    activeTab === tab.id
                      ? darkMode
                        ? "text-purple-400 border-b-2 border-purple-400 bg-gray-700"
                        : "text-indigo-600 border-b-2 border-indigo-600 bg-indigo-50"
                      : darkMode
                        ? "text-gray-400 hover:text-gray-200 hover:bg-gray-700"
                        : "text-gray-500 hover:text-gray-700 hover:bg-gray-50"
                  }`}
                >
                  <Icon className="h-4 w-4" />
                  {tab.label}
                </button>
              )
            })}
          </div>

          {/* Tab Content */}
          <div className="p-6">
            {/* Upload Tab */}
            {activeTab === "upload" && (
              <div className="space-y-6">
                {/* File Upload */}
                <div>
                  <h3 className={`text-lg font-semibold mb-4 ${darkMode ? "text-gray-100" : "text-gray-900"}`}>
                    Upload Excel File
                  </h3>
                  <div className="flex flex-wrap gap-4">
                    <input
                      type="file"
                      ref={fileInputRef}
                      onChange={handleFileSelect}
                      accept=".xlsx,.xls,.csv"
                      className="hidden"
                    />
                    <button
                      onClick={() => fileInputRef.current?.click()}
                      disabled={loading}
                      className={`flex items-center gap-2 px-6 py-3 rounded-lg disabled:opacity-50 disabled:cursor-not-allowed transition-colors ${
                        darkMode
                          ? "bg-purple-600 hover:bg-purple-700 text-white"
                          : "bg-indigo-600 hover:bg-indigo-700 text-white"
                      }`}
                    >
                      <Upload className="h-4 w-4" />
                      {loading ? "Uploading..." : "Choose File"}
                    </button>
                  </div>
                </div>

                {/* URL Upload */}
                <div>
                  <h3 className={`text-lg font-semibold mb-4 ${darkMode ? "text-gray-100" : "text-gray-900"}`}>
                    Upload from URL
                  </h3>
                  <div className="flex flex-wrap gap-4">
                    <input
                      type="text"
                      value={urlInput}
                      onChange={(e) => setUrlInput(e.target.value)}
                      placeholder="Enter Google Sheet URL..."
                      className={`flex-1 px-4 py-3 rounded-lg focus:ring-2 ${
                        darkMode
                          ? "bg-gray-700 border-gray-600 focus:ring-purple-500 text-white"
                          : "border border-gray-300 focus:ring-indigo-500 text-gray-900"
                      }`}
                    />
                    <button
                      onClick={() => uploadFromUrl(urlInput)}
                      disabled={loading || !urlInput}
                      className={`flex items-center gap-2 px-6 py-3 rounded-lg disabled:opacity-50 disabled:cursor-not-allowed transition-colors ${
                        darkMode
                          ? "bg-purple-600 hover:bg-purple-700 text-white"
                          : "bg-indigo-600 hover:bg-indigo-700 text-white"
                      }`}
                    >
                      <Link className="h-4 w-4" />
                      {loading ? "Loading..." : "Load from URL"}
                    </button>
                  </div>
                </div>

                {/* Uploaded Files List */}
                <div>
                  <h3 className={`text-lg font-semibold mb-4 ${darkMode ? "text-gray-100" : "text-gray-900"}`}>
                    Uploaded Files
                  </h3>
                  {uploadedFiles.length === 0 ? (
                    <p className={darkMode ? "text-gray-400" : "text-gray-500"}>No files uploaded yet.</p>
                  ) : (
                    <div className="space-y-2">
                      {uploadedFiles.map((file) => (
                        <div
                          key={file.id}
                          className={`flex items-center justify-between p-4 rounded-lg border ${
                            darkMode ? "bg-gray-700 border-gray-600" : "bg-gray-50 border-gray-200"
                          }`}
                        >
                          <div className="flex items-center gap-3">
                            <File className={`h-5 w-5 ${darkMode ? "text-purple-400" : "text-indigo-600"}`} />
                            <div>
                              <p className={`font-medium ${darkMode ? "text-gray-100" : "text-gray-900"}`}>
                                {file.name}
                              </p>
                              <p className={`text-sm ${darkMode ? "text-gray-400" : "text-gray-500"}`}>
                                {file.sheets.length} sheet(s) • Uploaded{" "}
                                {new Date(file.uploadedAt).toLocaleDateString()}
                              </p>
                            </div>
                          </div>
                          <button
                            onClick={() => deleteFile(file.id)}
                            className={`p-2 rounded-lg transition-colors ${
                              darkMode ? "text-red-400 hover:bg-red-900/30" : "text-red-600 hover:bg-red-50"
                            }`}
                            title="Delete file"
                          >
                            <Trash2 className="h-5 w-5" />
                          </button>
                        </div>
                      ))}
                    </div>
                  )}
                </div>
              </div>
            )}

            {/* Contacts Tab */}
            {activeTab === "contacts" && (
              <div>
                <div className="flex flex-wrap items-center gap-4 mb-6">
                  <button
                    onClick={() => fetchContacts()}
                    disabled={!selectedFile || !selectedSheet || loading}
                    className={`flex items-center gap-2 px-6 py-3 rounded-lg disabled:opacity-50 disabled:cursor-not-allowed transition-colors ${
                      darkMode
                        ? "bg-purple-600 hover:bg-purple-700 text-white"
                        : "bg-indigo-600 hover:bg-indigo-700 text-white"
                    }`}
                  >
                    <Users className="h-4 w-4" />
                    {loading ? "Loading..." : "Fetch Contacts"}
                  </button>

                  {data && data.contacts && (
                    <button
                      onClick={() => copyToClipboard(data.contacts.map((c) => c.email).join(", "))}
                      className={`flex items-center gap-2 px-4 py-3 rounded-lg transition-colors ${
                        darkMode
                          ? "bg-green-600 hover:bg-green-700 text-white"
                          : "bg-green-600 hover:bg-green-700 text-white"
                      }`}
                    >
                      <Copy className="h-4 w-4" />
                      Copy All Emails
                    </button>
                  )}
                </div>

                {error && (
                  <div
                    className={`${darkMode ? "bg-red-900/20 border-red-900 text-red-300" : "bg-red-50 border-red-200 text-red-700"} p-4 rounded-lg mb-4 flex items-center gap-3`}
                  >
                    <AlertCircle className="h-5 w-5" />
                    <p className="font-medium">Error: {error}</p>
                  </div>
                )}

                {loading && <p className={darkMode ? "text-gray-400" : "text-gray-500"}>Loading contacts...</p>}

                {!loading && data && data.contacts && (
                  <div>
                    <h3 className={`text-lg font-semibold mb-4 ${darkMode ? "text-gray-100" : "text-gray-900"}`}>
                      Contacts ({data.totalContacts})
                    </h3>
                    <div className="grid gap-2">
                      {data.contacts.length === 0 ? (
                        <p className={darkMode ? "text-gray-400" : "text-gray-500"}>No contacts found.</p>
                      ) : (
                        data.contacts.map((contact) => (
                          <div
                            key={contact.id}
                            className={`flex items-center justify-between p-3 rounded border ${
                              darkMode ? "bg-gray-800 border-gray-600" : "bg-white border-gray-200"
                            }`}
                          >
                            <div>
                              <p className={`font-medium ${darkMode ? "text-gray-100" : "text-gray-900"}`}>
                                {contact.name}
                              </p>
                              <p className={`text-sm ${darkMode ? "text-purple-300" : "text-gray-600"}`}>
                                {contact.email}
                              </p>
                            </div>
                            <div className="flex items-center gap-2">
                              {contact.isValidEmail ? (
                                <CheckCircle className="h-5 w-5 text-green-500" title="Valid Email" />
                              ) : ( 
                                <AlertCircle className="h-5 w-5 text-red-500" title="Invalid Email" />
                              )}
                              <button
                                onClick={() => copyToClipboard(contact.email)}
                                className={`${darkMode ? "text-purple-400 hover:text-purple-300" : "text-indigo-600 hover:text-indigo-800"}`}
                                title="Copy email"
                              >
                                <Copy className="h-4 w-4" />
                              </button>
                            </div>
                          </div>
                        ))
                      )}
                    </div>
                  </div>
                )}
              </div>
            )}

            {/* Emails Tab */}
            {activeTab === "emails" && (
              <div>
                <div className="flex flex-wrap items-center gap-4 mb-6">
                  <button
                    onClick={() => fetchEmails()}
                    disabled={!selectedFile || !selectedSheet || loading}
                    className={`flex items-center gap-2 px-6 py-3 rounded-lg disabled:opacity-50 disabled:cursor-not-allowed transition-colors ${
                      darkMode
                        ? "bg-purple-600 hover:bg-purple-700 text-white"
                        : "bg-indigo-600 hover:bg-indigo-700 text-white"
                    }`}
                  >
                    <Mail className="h-4 w-4" />
                    {loading ? "Loading..." : "Fetch Emails"}
                  </button>

                  {data && data.emails && (
                    <button
                      onClick={() => copyToClipboard(data.emails.join(", "))}
                      className={`flex items-center gap-2 px-4 py-3 rounded-lg transition-colors ${
                        darkMode
                          ? "bg-green-600 hover:bg-green-700 text-white"
                          : "bg-green-600 hover:bg-green-700 text-white"
                      }`}
                    >
                      <Copy className="h-4 w-4" />
                      Copy All Emails
                    </button>
                  )}
                </div>

                {error && (
                  <div
                    className={`${darkMode ? "bg-red-900/20 border-red-900 text-red-300" : "bg-red-50 border-red-200 text-red-700"} p-4 rounded-lg mb-4 flex items-center gap-3`}
                  >
                    <AlertCircle className="h-5 w-5" />
                    <p className="font-medium">Error: {error}</p>
                  </div>
                )}

                {loading && <p className={darkMode ? "text-gray-400" : "text-gray-500"}>Loading emails...</p>}

                {!loading && data && data.emails && (
                  <div>
                    <h3 className={`text-lg font-semibold mb-4 ${darkMode ? "text-gray-100" : "text-gray-900"}`}>
                      Emails ({data.totalEmails})
                    </h3>
                    <div
                      className={`p-4 rounded-lg ${
                        darkMode ? "bg-gray-800 border border-gray-700" : "bg-gray-100 border border-gray-200"
                      }`}
                    >
                      <div className="grid gap-2">
                        {data.emails.map((email, index) => (
                          <div
                            key={index}
                            className={`flex items-center justify-between p-3 rounded border ${
                              darkMode ? "bg-gray-800 border-gray-600" : "bg-white border-gray-200"
                            }`}
                          >
                            <span className={`font-mono text-sm ${darkMode ? "text-purple-300" : ""}`}>{email}</span>
                            <button
                              onClick={() => copyToClipboard(email)}
                              className={`${darkMode ? "text-purple-400 hover:text-purple-300" : "text-indigo-600 hover:text-indigo-800"}`}
                              title="Copy email"
                            >
                              <Copy className="h-4 w-4" />
                            </button>
                          </div>
                        ))}
                      </div>
                    </div>
                  </div>
                )}
              </div>
            )}

            {/* Sheet Data Tab */}
            {activeTab === "sheet" && (
              <div>
                <div className="flex flex-wrap items-center gap-4 mb-6">
                  <button
                    onClick={() => fetchFileData()}
                    disabled={!selectedFile || !selectedSheet || loading}
                    className={`flex items-center gap-2 px-6 py-3 rounded-lg disabled:opacity-50 disabled:cursor-not-allowed transition-colors ${
                      darkMode
                        ? "bg-purple-600 hover:bg-purple-700 text-white"
                        : "bg-indigo-600 hover:bg-indigo-700 text-white"
                    }`}
                  >
                    <FileSpreadsheet className="h-4 w-4" />
                    {loading ? "Loading..." : "Fetch Sheet Data"}
                  </button>
                  {data && data.data && (
                    <button
                      onClick={downloadData}
                      className={`flex items-center gap-2 px-4 py-3 rounded-lg transition-colors ${
                        darkMode
                          ? "bg-blue-600 hover:bg-blue-700 text-white"
                          : "bg-blue-600 hover:bg-blue-700 text-white"
                      }`}
                    >
                      <Download className="h-4 w-4" />
                      Download JSON
                    </button>
                  )}
                </div>

                {error && (
                  <div
                    className={`${darkMode ? "bg-red-900/20 border-red-900 text-red-300" : "bg-red-50 border-red-200 text-red-700"} p-4 rounded-lg mb-4 flex items-center gap-3`}
                  >
                    <AlertCircle className="h-5 w-5" />
                    <p className="font-medium">Error: {error}</p>
                  </div>
                )}

                {loading && <p className={darkMode ? "text-gray-400" : "text-gray-500"}>Loading sheet data...</p>}

                {!loading && data && data.data && (
                  <div>
                    <h3 className={`text-lg font-semibold mb-4 ${darkMode ? "text-gray-100" : "text-gray-900"}`}>
                      Sheet Data ({data.rowCount} rows, {data.columnCount} columns)
                    </h3>
                    <div
                      className={`p-4 rounded-lg overflow-auto max-h-96 ${darkMode ? "bg-gray-800 border border-gray-700" : "bg-gray-100 border border-gray-200"}`}
                    >
                      <pre className="whitespace-pre-wrap font-mono text-sm">
                        {JSON.stringify(data.data, null, 2)}
                      </pre>
                    </div>
                  </div>
                )}
              </div>
            )}

            {/* Search Tab */}
            {activeTab === "search" && (
              <div>
                <div className="flex flex-wrap items-center gap-4 mb-6">
                  <div className="flex-1">
                    <label className={`block text-sm font-medium mb-1 ${darkMode ? "text-gray-300" : "text-gray-700"}`}>
                      Search Query
                    </label>
                    <input
                      type="text"
                      value={searchQuery}
                      onChange={(e) => setSearchQuery(e.target.value)}
                      onKeyPress={handleKeyPress}
                      placeholder="Search by name or email..."
                      className={`w-full px-4 py-2 rounded-lg focus:ring-2 ${
                        darkMode
                          ? "bg-gray-700 border-gray-600 focus:ring-purple-500 text-white"
                          : "border border-gray-300 focus:ring-indigo-500 text-gray-900"
                      }`}
                    />
                  </div>
                  <button
                    onClick={searchContacts}
                    disabled={!searchQuery || !selectedFile || !selectedSheet || loading}
                    className={`flex items-center gap-2 px-6 py-3 rounded-lg disabled:opacity-50 disabled:cursor-not-allowed transition-colors ${
                      darkMode
                        ? "bg-purple-600 hover:bg-purple-700 text-white"
                        : "bg-indigo-600 hover:bg-indigo-700 text-white"
                    }`}
                  >
                    <Search className="h-4 w-4" />
                    {loading ? "Searching..." : "Search Contacts"}
                  </button>
                </div>

                {error && (
                  <div
                    className={`${darkMode ? "bg-red-900/20 border-red-900 text-red-300" : "bg-red-50 border-red-200 text-red-700"} p-4 rounded-lg mb-4 flex items-center gap-3`}
                  >
                    <AlertCircle className="h-5 w-5" />
                    <p className="font-medium">Error: {error}</p>
                  </div>
                )}

                {loading && <p className={darkMode ? "text-gray-400" : "text-gray-500"}>Searching...</p>}

                {!loading && data && data.results && (
                  <div>
                    <h3 className={`text-lg font-semibold mb-4 ${darkMode ? "text-gray-100" : "text-gray-900"}`}>
                      Search Results ({data.totalFound})
                    </h3>
                    <div className="grid gap-2">
                      {data.results.length === 0 ? (
                        <p className={darkMode ? "text-gray-400" : "text-gray-500"}>No results found.</p>
                      ) : (
                        data.results.map((contact) => (
                          <div
                            key={contact.id}
                            className={`flex items-center justify-between p-3 rounded border ${
                              darkMode ? "bg-gray-800 border-gray-600" : "bg-white border-gray-200"
                            }`}
                          >
                            <div>
                              <p className={`font-medium ${darkMode ? "text-gray-100" : "text-gray-900"}`}>
                                {contact.name}
                              </p>
                              <p className={`text-sm ${darkMode ? "text-purple-300" : "text-gray-600"}`}>
                                {contact.email}
                              </p>
                            </div>
                            <div className="flex items-center gap-2">
                              {contact.isValidEmail ? (
                                <CheckCircle className="h-5 w-5 text-green-500" title="Valid Email" />
                              ) : (
                                <AlertCircle className="h-5 w-5 text-red-500" title="Invalid Email" />
                              )}
                              <button
                                onClick={() => copyToClipboard(contact.email)}
                                className={`${darkMode ? "text-purple-400 hover:text-purple-300" : "text-indigo-600 hover:text-indigo-800"}`}
                                title="Copy email"
                              >
                                <Copy className="h-4 w-4" />
                              </button>
                            </div>
                          </div>
                        ))
                      )}
                    </div>
                  </div>
                )}
              </div>
            )}

            {/* Send Email Tab */}
            {activeTab === "sendmail" && (
              <div>
                <div className="mb-6">
                  <h2 className={`text-xl font-bold mb-2 ${darkMode ? "text-gray-100" : "text-gray-900"}`}>
                    Send Email to All Valid Emails
                  </h2>
                  {/* Template Selector */}
                  <div className="mb-4">
                    <label className={`block text-sm font-medium mb-1 ${darkMode ? "text-gray-300" : "text-gray-700"}`}>
                      Choose Template
                    </label>
                    <div className="flex flex-wrap gap-2 w-full">
                      <select
                        value={selectedTemplate}
                        onChange={handleTemplateChange}
                        className={`px-3 py-2 rounded-md flex-1 ${
                          darkMode ? "bg-gray-700 text-white border border-gray-600" : "bg-gray-700 text-white"
                        }`}
                      >
                        <option value="">Select Template</option>
                        <option value="welcome">Welcome Email</option>
                        <option value="newsletter">Newsletter</option>
                        <option value="certificate">Certificate Template</option>
                      </select>
                      <button
                        onClick={() => selectedTemplate && loadTemplateInEditor(selectedTemplate)}
                        disabled={!selectedTemplate}
                        className={`px-4 py-2 rounded-md disabled:cursor-not-allowed ${
                          darkMode
                            ? "bg-blue-600 hover:bg-blue-700 text-white disabled:bg-gray-600"
                            : "bg-blue-600 hover:bg-blue-700 text-white disabled:bg-gray-600"
                        }`}
                      >
                        Load Template
                      </button>
                    </div>
                  </div>
                  <div className="mb-4">
                    <label className={`block text-sm font-medium mb-1 ${darkMode ? "text-gray-300" : "text-gray-700"}`}>
                      Subject
                    </label>
                    <input
                      type="text"
                      value={emailSubject}
                      onChange={(e) => setEmailSubject(e.target.value)}
                      className={`w-full px-4 py-2 rounded-lg ${
                        darkMode ? "bg-gray-700 border-gray-600 text-white" : "border border-gray-300"
                      }`}
                      placeholder="Enter email subject"
                    />
                  </div>
                  {/* Email Editor */}
                  <div className="mb-4">
                    <label className={`block text-sm font-medium mb-1 ${darkMode ? "text-gray-300" : "text-gray-700"}`}>
                      Body
                    </label>
                    <div
                      className={`w-full rounded-lg overflow-hidden ${darkMode ? "border border-gray-600" : "border border-gray-300"}`}
                      style={{
                        height: "600px",
                        width: "100%",
                        position: "relative",
                        display: "flex",
                        flexDirection: "column"
                      }}
                    >
                      {activeTab === "sendmail" && (
                        <EmailEditor
                          ref={emailEditorRef}
                          onLoad={handleEditorLoad}
                          onReady={() => console.log("Editor is ready")}
                          style={{
                            height: "100%",
                            width: "100%",
                            display: "flex",
                            flexDirection: "column"
                          }}
                          options={{
                            displayMode: "email",
                            safeHtml: true,
                            features: {
                              preheaderText: false,
                              textEditor: {
                                tables: true,
                                emojis: true
                              }
                            },
                            tools: {
                              image: { enabled: true },
                              button: { enabled: true },
                              text: { enabled: true },
                              form: { enabled: true },
                              divider: { enabled: true },
                              social: { enabled: true },
                              timer: { enabled: false },
                              video: { enabled: false },
                              menu: { enabled: false }
                            },
                            appearance: {
                              theme: darkMode ? "dark" : "light",
                              panels: {
                                tools: {
                                  dock: "left",
                                },
                              },
                            },
                            customCSS: [
                              `
                              #editor-container {
                                height: 100% !important;
                                min-height: 600px !important;
                                width: 100% !important;
                                display: flex !important;
                                flex-direction: column !important;
                              }
                              .unlayer-wrapper, .unlayer-container {
                                height: 100% !important;
                                width: 100% !important;
                                display: flex !important;
                                flex-direction: column !important;
                              }
                              .unlayer-editor {
                                flex: 1 !important;
                                display: flex !important;
                                flex-direction: column !important;
                              }
                              `,
                            ],
                            mergeTags: [
                              {
                                name: 'User',
                                mergeTags: [
                                  {
                                    name: 'Name',
                                    value: '{{name}}',
                                  },
                                  {
                                    name: 'Email',
                                    value: '{{email}}',
                                  }
                                ]
                              }
                            ]
                          }}
                        />
                      )}
                    </div>
                  </div>
                  <button
                    onClick={async () => {
                      if (emailEditorRef.current && emailEditorRef.current.editor) {
                        emailEditorRef.current.editor.exportHtml((data) => {
                          setEmailBody(data.html)
                          sendEmails()
                        })
                      } else {
                        sendEmails()
                      }
                    }}
                    disabled={!emailSubject || !selectedFile || !selectedSheet || loading}
                    className={`px-6 py-3 rounded-lg disabled:opacity-50 disabled:cursor-not-allowed transition-colors ${
                      darkMode
                        ? "bg-purple-600 hover:bg-purple-700 text-white"
                        : "bg-indigo-600 hover:bg-indigo-700 text-white"
                    }`}
                  >
                    {loading ? "Sending..." : "Send Email"}
                  </button>
                </div>
                {emailSendResult && (
                  <div
                    className={`${emailSendResult.failed > 0
                        ? darkMode
                          ? "bg-red-900/20 border-red-900 text-red-300"
                          : "bg-red-50 border-red-200 text-red-700"
                        : darkMode
                          ? "bg-green-900/20 border-green-900 text-green-300"
                          : "bg-green-50 border-green-200 text-green-700"
                      } p-4 rounded-lg mt-4 flex items-center gap-3`}
                  >
                    {emailSendResult.failed > 0 ? (
                      <AlertCircle className="h-5 w-5" />
                    ) : (
                      <CheckCircle className="h-5 w-5" />
                    )}
                    <p className="font-medium">
                      {emailSendResult.message} (Sent: {emailSendResult.sent}, Failed:{" "}
                      {emailSendResult.failed})
                    </p>
                  </div>
                )}
                {emailSendResult && emailSendResult.errors.length > 0 && (
                  <div
                    className={`${darkMode ? "bg-gray-800 border border-gray-700" : "bg-gray-100 border border-gray-200"} p-4 rounded-lg mt-4`}
                  >
                    <h4 className={`text-md font-semibold mb-2 ${darkMode ? "text-gray-100" : "text-gray-900"}`}>
                      Details of failed emails:
                    </h4>
                    <ul className="list-disc list-inside">
                      {emailSendResult.errors.map((err, index) => (
                        <li key={index} className={darkMode ? "text-gray-300" : "text-gray-700"}>
                          {err.email}: {err.error}
                        </li>
                      ))}
                    </ul>
                  </div>
                )}
              </div>
            )}
          </div>
        </div>
      </div>
    </div>
  )
}

export default XlsxUploadFrontend
