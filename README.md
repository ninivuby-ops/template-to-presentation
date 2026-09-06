# Template To Presentation

Build a complete web application that allows users to upload a PowerPoint (.pptx) template and automatically generates a fully populated presentation using AI. 

Requirements:

1. The app must strictly use ONLY the uploaded template file. 

2. Do not change the slide sequence, section headings, prescribed pointers, or overall template structure. 

3. Do not add or remove slides. 

4. Fill in the content strictly within the placeholders provided in the template. 

5. Maintain the fonts, colors, and layout exactly as defined in the uploaded template. 

6. The output must be a fully populated PPTX file that adheres 100% to the uploaded format.

Technical specifications:

- Frontend: Simple upload form (HTML/CSS/JavaScript or React).

- Backend: Python (Flask or Django).

- File handling: Use `python-pptx` to parse and modify the uploaded template.

- AI integration: Use OpenAI API (or similar) to generate text content for placeholders.

- Workflow: 

   a. User uploads template.  

   b. Backend extracts placeholders.  

   c. AI generates content based on headings/pointers.  

   d. Backend inserts content back into template.  

   e. User downloads final PPTX.  

Deliverables:

- Full backend code (Flask/Django).  

- Frontend upload interface.  

- AI prompt integration.  

- Example API call structure.  

- End-to-end workflow ensuring strict adherence to template rules. with genearted image s, architecture, diagrams.

This project was built with [Lovable](https://lovable.dev).

## Build with Lovable

Continue developing this project in the [Lovable editor](https://lovable.dev/projects/376fa59e-4fc1-462a-9665-569252960e0f).

- **Ship faster**: describe what you want to build and Lovable handles the code.
- **Stay in sync**: every change made in Lovable is committed straight to this repository.
- **Full ownership**: this code is yours. Push to `main` on GitHub and your changes sync back into Lovable, ready for your next prompt.

## Development

Prefer working locally? You need Node.js and npm — [install with nvm](https://github.com/nvm-sh/nvm#installing-and-updating).

```sh
git clone <this-repository-url>
cd <repository-name>
npm i
npm run dev
```
