FROM ghcr.io/astral-sh/uv:python3.13-trixie

COPY .python-version .
COPY pyproject.toml .
COPY uv.lock .

RUN uv sync --locked

# create user with a home directory
ARG NB_USER=jovyan
ARG NB_UID=1000
ENV USER=${NB_USER}
ENV NB_UID=${NB_UID}
ENV HOME=/home/${NB_USER}

RUN adduser --disabled-password \
    --gecos "Default user" \
    --uid ${NB_UID} \
    ${NB_USER}

# Make sure the contents of our repo are in ${HOME}
COPY . ${HOME}
USER root
RUN chown -R ${NB_UID} ${HOME}
USER ${NB_USER}
WORKDIR ${HOME}

# Set up the virtual environment path
ENV VIRTUAL_ENV=/root/.venv
ENV PATH="/.venv/bin:$PATH"
ENV PATH=/root/.local/bin:$PATH