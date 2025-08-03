---
applyTo: "**"
---
# Project general coding standards
## Response language
- Use Vietnamese for all code comments, documentation, and commit messages

## readme.md
- Tất cả các tệp README.md phải được viết bằng tiếng Việt, giải thích rõ ràng về cách sử dụng, cấu hình và triển khai dự án.
- Chỉ tạo 1 file README.md duy nhất trong thư mục gốc của dự án, không tạo các tệp README.md khác trong các thư mục con.

## React và JavaScript
- Sử dụng React cho giao diện người dùng, không sử dụng các framework khác như Angular hoặc Vue.js.
- Sử dụng JavaScript ES6+ cho mã nguồn, không sử dụng các phiên bản cũ hơn.
- Sử dụng `const` và `let` thay vì `var` để khai báo biến.
- Sử dụng arrow functions thay vì function expressions truyền thống.
- Sử dụng destructuring để truy cập các thuộc tính của đối tượng.
- Sử dụng template literals thay vì chuỗi nối.              
- Sử dụng `async/await` thay vì callback hoặc Promise chaining để xử lý bất đồng bộ.
- Sử dụng `import` và `export` thay vì `require` và `module.exports` để quản lý module.
- Sử dụng `class` thay vì prototype-based inheritance cho các component React.
- Sử dụng `useState`, `useEffect` và các hook khác để quản lý state và lifecycle trong component function.
- Sử dụng `PropTypes` hoặc TypeScript để xác định kiểu dữ liệu của props trong component.
- Sử dụng `styled-components` để tạo kiểu cho các component React.
- Sử dụng `axios` hoặc `fetch` để thực hiện các yêu cầu HTTP, không sử dụng jQuery.
- Sử dụng `ESLint` và `Prettier` để kiểm tra và định dạng mã nguồn.
- Sử dụng `Jest` và `React Testing Library` để viết và chạy các bài kiểm tra đơn vị cho component React.
- Sử dụng `Redux` hoặc `Context API` để quản lý state toàn cục, không sử dụng các thư viện khác như MobX hoặc Vuex.
- Sử dụng `React Router` để quản lý routing trong ứng dụng, không sử dụng các thư viện khác như Vue Router hoặc Angular Router.
- Sử dụng `CSS Modules` hoặc `styled-components` để tạo kiểu cho các component, không sử dụng các thư viện CSS khác như Bootstrap hoặc Material UI. 
- Sử dụng `npm` hoặc `yarn` để quản lý các package, không sử dụng các trình quản lý package khác như Bower hoặc Composer.
