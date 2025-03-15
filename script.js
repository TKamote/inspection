document.addEventListener("DOMContentLoaded", () => {
  const cardsContainer = document.getElementById("cardsContainer");
  const addCardBtn = document.getElementById("addCardBtn");
  const uploadBtn = document.getElementById("uploadBtn");
  const resetBtn = document.getElementById("resetBtn");

  let cardCounter = 1;

  // Handle image upload and preview
  function handleImageUpload(input, previewDiv) {
    const file = input.files[0];
    if (file) {
      const reader = new FileReader();
      reader.onload = function (e) {
        previewDiv.style.backgroundImage = `url(${e.target.result})`;
        previewDiv.parentElement.querySelector(
          ".upload-overlay"
        ).style.display = "none";
      };
      reader.readAsDataURL(file);
    }
  }

  // Create a new card
  function createCard() {
    cardCounter++;
    const cardId = cardCounter;

    const cardDiv = document.createElement("div");
    cardDiv.className = "card";
    cardDiv.dataset.cardId = cardId;

    cardDiv.innerHTML = `
            <div class="input-group">
                <label for="serialNumber${cardId}">Serial Number:</label>
                <input type="text" id="serialNumber${cardId}" class="serial-number" required>
            </div>
            
            <div class="input-group">
                <label for="location${cardId}">Location:</label>
                <input type="text" id="location${cardId}" class="location" required>
            </div>
            
            <div class="input-group">
                <label for="comments${cardId}">Comments:</label>
                <textarea id="comments${cardId}" class="comments" rows="3"></textarea>
            </div>
            
            <div class="input-group">
                <label for="image${cardId}">Image:</label>
                <div class="image-upload-container">
                    <input type="file" id="image${cardId}" class="image-input" accept="image/*" capture="environment">
                    <div class="image-preview"></div>
                    <div class="upload-overlay">
                        <i class="fas fa-camera"></i>
                        <span>Take Photo or Upload</span>
                    </div>
                </div>
            </div>

            <button class="delete-btn">
                <i class="fas fa-trash"></i>
            </button>
        `;

    cardsContainer.appendChild(cardDiv);

    // Set up image upload handler for the new card
    const imageInput = cardDiv.querySelector(".image-input");
    const imagePreview = cardDiv.querySelector(".image-preview");
    imageInput.addEventListener("change", () =>
      handleImageUpload(imageInput, imagePreview)
    );

    return cardDiv;
  }

  // Handle delete button clicks
  cardsContainer.addEventListener("click", (e) => {
    if (e.target.closest(".delete-btn")) {
      const card = e.target.closest(".card");
      const cards = document.querySelectorAll(".card");

      if (cards.length > 1) {
        card.remove();
      }
    }
  });

  // Set up image upload handler for the initial card
  const initialImageInput = document.querySelector(".image-input");
  const initialImagePreview = document.querySelector(".image-preview");
  initialImageInput.addEventListener("change", () =>
    handleImageUpload(initialImageInput, initialImagePreview)
  );

  // Add new card button handler
  addCardBtn.addEventListener("click", createCard);

  // Reset button handler
  resetBtn.addEventListener("click", () => {
    const confirmation = confirm(
      "Are you sure you want to reset? This will delete all cards except the first one."
    );
    if (confirmation) {
      const cards = Array.from(document.querySelectorAll(".card"));
      cards.slice(1).forEach((card) => card.remove());

      // Reset the first card
      const firstCard = cards[0];
      firstCard
        .querySelectorAll('input[type="text"], textarea')
        .forEach((input) => (input.value = ""));
      const imagePreview = firstCard.querySelector(".image-preview");
      imagePreview.style.backgroundImage = "";
      firstCard.querySelector(".upload-overlay").style.display = "flex";

      cardCounter = 1;
    }
  });

  // Upload button handler
  uploadBtn.addEventListener("click", async () => {
    const cards = Array.from(document.querySelectorAll(".card"));

    // Validate that all required fields are filled
    const isValid = cards.every((card) => {
      const requiredInputs = card.querySelectorAll("[required]");
      return Array.from(requiredInputs).every(
        (input) => input.value.trim() !== ""
      );
    });

    if (!isValid) {
      alert(
        "Please fill in all required fields (Serial Number and Location) before uploading."
      );
      return;
    }

    // Create a new Document
    const doc = new docx.Document({
      sections: [
        {
          properties: {
            page: {
              margin: {
                top: 1000,
                right: 1000,
                bottom: 1000,
                left: 1000,
              },
            },
          },
          children: await generateDocumentContent(cards),
        },
      ],
    });

    // Generate and download the document
    const blob = await docx.Packer.toBlob(doc);
    const url = window.URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = "cards-export.docx";
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    window.URL.revokeObjectURL(url);
  });

  // Generate document content
  async function generateDocumentContent(cards) {
    const children = [];
    const cardsPerRow = 2;

    for (let i = 0; i < cards.length; i += cardsPerRow * 2) {
      const rowCards = cards.slice(i, i + cardsPerRow * 2);
      const tableRows = [];

      for (let j = 0; j < rowCards.length; j += cardsPerRow) {
        const row = rowCards.slice(j, j + cardsPerRow);
        const tableCells = await Promise.all(
          row.map(async (card) => {
            const serialNumber = card.querySelector(".serial-number").value;
            const location = card.querySelector(".location").value;
            const comments = card.querySelector(".comments").value;
            const imagePreview = card.querySelector(".image-preview");
            const imageUrl = imagePreview.style.backgroundImage.slice(5, -2);

            let imageData;
            if (imageUrl) {
              try {
                const response = await fetch(imageUrl);
                const blob = await response.blob();
                imageData = await new Promise((resolve) => {
                  const reader = new FileReader();
                  reader.onloadend = () => resolve(reader.result.split(",")[1]);
                  reader.readAsDataURL(blob);
                });
              } catch (error) {
                console.error("Error processing image:", error);
              }
            }

            return new docx.TableCell({
              width: {
                size: 4500,
                type: docx.WidthType.DXA,
              },
              margins: {
                top: 100,
                bottom: 100,
                left: 100,
                right: 100,
              },
              children: [
                new docx.Paragraph({
                  children: [
                    new docx.TextRun({
                      text: `Serial Number: ${serialNumber}`,
                      bold: true,
                    }),
                  ],
                }),
                new docx.Paragraph({
                  children: [
                    new docx.TextRun({
                      text: `Location: ${location}`,
                    }),
                  ],
                }),
                new docx.Paragraph({
                  children: [
                    new docx.TextRun({
                      text: `Comments: ${comments}`,
                    }),
                  ],
                }),
                imageData &&
                  new docx.Paragraph({
                    children: [
                      new docx.ImageRun({
                        data: imageData,
                        transformation: {
                          width: 200,
                          height: 200,
                        },
                      }),
                    ],
                  }),
              ].filter(Boolean),
            });
          })
        );

        while (tableCells.length < cardsPerRow) {
          tableCells.push(
            new docx.TableCell({
              width: {
                size: 4500,
                type: docx.WidthType.DXA,
              },
              children: [new docx.Paragraph({})],
            })
          );
        }

        tableRows.push(
          new docx.TableRow({
            children: tableCells,
          })
        );
      }

      children.push(
        new docx.Table({
          width: {
            size: 9000,
            type: docx.WidthType.DXA,
          },
          rows: tableRows,
        }),
        new docx.Paragraph({})
      );
    }

    return children;
  }
});
